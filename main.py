import os
import io
import time
import gzip
import shutil
import zipfile
import logging
import traceback
from typing import Generator, List
from concurrent.futures import ThreadPoolExecutor
import httpx
from fastapi import FastAPI, HTTPException, Response
from fastapi.responses import StreamingResponse, HTMLResponse, FileResponse
from fastapi.background import BackgroundTasks
from pydantic import BaseModel
import clickhouse_connect
import xlsxwriter
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s - %(message)s",
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger("ClickhouseExporter")

app = FastAPI(title="ClickHouse Query Executor & Excel Exporter")

# Read credentials from environment
CLICKHOUSE_HOST = os.getenv("CLICKHOUSE_HOST", "localhost")
CLICKHOUSE_PORT = int(os.getenv("CLICKHOUSE_PORT", "8123"))
CLICKHOUSE_USER = os.getenv("CLICKHOUSE_USER", "default")
CLICKHOUSE_PASSWORD = os.getenv("CLICKHOUSE_PASSWORD", "1144")

# Temporary directory for zip creations
TEMP_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "temp_exports")
os.makedirs(TEMP_DIR, exist_ok=True)

class GzipStreamer:
    """Helper class to compress chunks on the fly into a gzip stream."""
    def __init__(self, compresslevel=6):
        self.buffer = io.BytesIO()
        self.zipfile = gzip.GzipFile(mode='wb', fileobj=self.buffer, compresslevel=compresslevel)

    def compress(self, data: bytes) -> bytes:
        self.zipfile.write(data)
        self.zipfile.flush()
        compressed = self.buffer.getvalue()
        self.buffer.seek(0)
        self.buffer.truncate(0)
        return compressed

    def close(self) -> bytes:
        self.zipfile.close()
        compressed = self.buffer.getvalue()
        self.buffer.seek(0)
        self.buffer.truncate(0)
        return compressed

def get_clickhouse_client():
    try:
        logger.info(f"Connecting to ClickHouse at {CLICKHOUSE_HOST}:{CLICKHOUSE_PORT}...")
        return clickhouse_connect.get_client(
            host=CLICKHOUSE_HOST,
            port=CLICKHOUSE_PORT,
            username=CLICKHOUSE_USER,
            password=CLICKHOUSE_PASSWORD,
            connect_timeout=15
        )
    except Exception as e:
        logger.error(f"ClickHouse connection failed: {str(e)}\n{traceback.format_exc()}")
        raise HTTPException(status_code=500, detail=f"Failed to connect to ClickHouse: {str(e)}")

class QueryRequest(BaseModel):
    query: str
    filename: str = "query_export"

@app.get("/", response_class=HTMLResponse)
def read_index():
    index_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "index.html")
    if os.path.exists(index_path):
        with open(index_path, "r", encoding="utf-8") as f:
            return f.read()
    return "<h3>index.html not found</h3>"

def cleanup_file(path: str):
    try:
        logger.info(f"Cleaning up temporary file/folder: {path}")
        if os.path.isdir(path):
            shutil.rmtree(path)
        elif os.path.exists(path):
            os.remove(path)
        logger.info(f"Cleanup successful for {path}")
    except Exception as e:
        logger.warning(f"Error during cleanup of {path}: {str(e)}")

def write_excel_chunk(
    host: str,
    port: int,
    user: str,
    password: str,
    raw_query: str,
    limit: int,
    offset: int,
    file_path: str,
    col_names: List[str],
    is_string_col: List[bool]
) -> str:
    client = None
    try:
        logger.info(f"Thread started for offset {offset} writing to {file_path}")
        client = clickhouse_connect.get_client(
            host=host,
            port=port,
            username=user,
            password=password,
            connect_timeout=30
        )
        
        chunk_query = f"SELECT * FROM ({raw_query}) LIMIT {limit} OFFSET {offset}"
        result = client.query(chunk_query)
        
        workbook = xlsxwriter.Workbook(file_path, {'constant_memory': True, 'in_memory': False})
        worksheet = workbook.add_worksheet("Query Results")
        text_format = workbook.add_format({'num_format': '@'})
        
        for col_idx, col_name in enumerate(col_names):
            worksheet.write(0, col_idx, col_name)
            
        for row_idx, row in enumerate(result.result_rows, start=1):
            for col_idx, val in enumerate(row):
                if val is None:
                    worksheet.write_blank(row_idx, col_idx, None)
                elif is_string_col[col_idx]:
                    worksheet.write_string(row_idx, col_idx, str(val), text_format)
                else:
                    if isinstance(val, (int, float)):
                        worksheet.write_number(row_idx, col_idx, val)
                    elif isinstance(val, bool):
                        worksheet.write_boolean(row_idx, col_idx, val)
                    else:
                        worksheet.write(row_idx, col_idx, str(val))
                        
        workbook.close()
        logger.info(f"Thread completed for offset {offset}. File written: {file_path}")
    except Exception as e:
        logger.error(f"Thread failed for offset {offset}. Error: {str(e)}\n{traceback.format_exc()}")
        raise e
    finally:
        if client:
            try:
                client.close()
            except Exception:
                pass
    return file_path

def validate_select_query(query: str):
    import re
    # Remove SQL comments -- or /* */
    query_no_comments = re.sub(r'--.*$', '', query, flags=re.MULTILINE)
    query_no_comments = re.sub(r'/\*.*?\*/', '', query_no_comments, flags=re.DOTALL)
    query_stripped = query_no_comments.strip().upper()
    
    if not (query_stripped.startswith("SELECT") or query_stripped.startswith("WITH")):
        raise HTTPException(status_code=400, detail="Only SELECT or WITH queries are allowed.")

EXPORT_TASKS = {}

def cleanup_old_files():
    now = time.time()
    if os.path.exists(TEMP_DIR):
        for item in os.listdir(TEMP_DIR):
            item_path = os.path.join(TEMP_DIR, item)
            try:
                if os.path.getmtime(item_path) < now - 600:
                    if os.path.isdir(item_path):
                        shutil.rmtree(item_path)
                    else:
                        os.remove(item_path)
            except Exception as e:
                logger.warning(f"Error during scheduled cleanup of {item_path}: {e}")

@app.get("/api/exports")
def list_exports():
    cleanup_old_files()
    exports = []
    if os.path.exists(TEMP_DIR):
        for entry in os.listdir(TEMP_DIR):
            entry_path = os.path.join(TEMP_DIR, entry)
            if os.path.isdir(entry_path):
                for f in os.listdir(entry_path):
                    f_path = os.path.join(entry_path, f)
                    if os.path.isfile(f_path):
                        stat = os.stat(f_path)
                        exports.append({
                            "filename": f,
                            "folder": entry,
                            "size_mb": round(stat.st_size / (1024 * 1024), 2),
                            "created_at": stat.st_mtime
                        })
    exports.sort(key=lambda x: x["created_at"], reverse=True)
    return exports

@app.delete("/api/exports/delete/{folder}/{filename}")
def delete_export(folder: str, filename: str):
    folder = "".join(c for c in folder if c.isalnum() or c in ('_', '-'))
    filename = "".join(c for c in filename if c.isalnum() or c in ('_', '-', '.'))
    
    file_dir = os.path.join(TEMP_DIR, folder)
    file_path = os.path.join(file_dir, filename)
    
    if os.path.exists(file_path):
        try:
            os.remove(file_path)
            if os.path.exists(file_dir) and not os.listdir(file_dir):
                shutil.rmtree(file_dir)
            return {"status": "deleted"}
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Failed to delete file: {str(e)}")
    raise HTTPException(status_code=404, detail="File not found")

@app.get("/api/exports/download/{folder}/{filename}")
def download_cached_file(folder: str, filename: str):
    folder = "".join(c for c in folder if c.isalnum() or c in ('_', '-'))
    filename = "".join(c for c in filename if c.isalnum() or c in ('_', '-', '.'))
    
    file_path = os.path.join(TEMP_DIR, folder, filename)
    if os.path.exists(file_path) and os.path.isfile(file_path):
        media_type = "application/octet-stream"
        if filename.endswith(".xlsx"):
            media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        elif filename.endswith(".csv"):
            media_type = "text/csv"
        elif filename.endswith(".tsv"):
            media_type = "text/tab-separated-values"
        elif filename.endswith(".gz"):
            media_type = "application/gzip"
            
        headers = {
            'Content-Disposition': f'attachment; filename="{filename}"'
        }
        return FileResponse(file_path, media_type=media_type, headers=headers)
    raise HTTPException(status_code=404, detail="File not found or expired")

@app.get("/api/exports/status/{task_id}")
def get_export_status(task_id: str):
    if task_id not in EXPORT_TASKS:
        raise HTTPException(status_code=404, detail="Task not found")
        
    task = EXPORT_TASKS[task_id]
    status = task["status"]
    processed = task["processed_rows"]
    total = task["total_rows"]
    
    elapsed = time.time() - task["start_time"] if status == "processing" else task["elapsed"]
    
    remaining = 0
    if status == "processing" and processed > 5000 and total > 0:
        rate = processed / elapsed
        if rate > 0:
            remaining = int((total - processed) / rate)
            
    return {
        "status": status,
        "processed_rows": processed,
        "total_rows": total,
        "elapsed_seconds": int(elapsed),
        "estimated_remaining_seconds": remaining,
        "error": task["error"],
        "download_url": f"/api/exports/download/{task['folder']}/{task['filename']}" if status == "completed" else None
    }

@app.post("/api/exports/cancel/{task_id}")
def cancel_export_task(task_id: str):
    if task_id in EXPORT_TASKS:
        EXPORT_TASKS[task_id]["abort_requested"] = True
        EXPORT_TASKS[task_id]["status"] = "cancelled"
        return {"status": "cancelled"}
    raise HTTPException(status_code=404, detail="Task not found")

def async_export_worker(task_id: str, raw_query: str, filename: str, format: str, timestamp: str):
    task = EXPORT_TASKS[task_id]
    client = None
    task_dir = None
    try:
        client = get_clickhouse_client()
        
        logger.info(f"[{task_id}] Executing Clickhouse query for export...")
        result = client.query(raw_query)
        
        column_names = result.column_names
        is_string_col = ["string" in str(t).lower() or "uuid" in str(t).lower() for t in result.column_types]
        
        rows = result.result_rows
        task["total_rows"] = len(rows)
        
        task_dir = os.path.join(TEMP_DIR, task_id)
        os.makedirs(task_dir, exist_ok=True)
        final_filename = f"{filename}_{timestamp}.{format}"
        file_path = os.path.join(task_dir, final_filename)
        
        task["file_path"] = file_path
        task["filename"] = final_filename
        task["folder"] = task_id
        
        if format == "xlsx":
            workbook = xlsxwriter.Workbook(file_path, {'constant_memory': True, 'in_memory': False})
            
            header_format = workbook.add_format({
                'bold': True, 'border': 1, 'border_color': '#000000', 'bg_color': '#e1e5eb'
            })
            cell_format = workbook.add_format({'border': 1, 'border_color': '#000000'})
            text_format = workbook.add_format({'num_format': '@', 'border': 1, 'border_color': '#000000'})
            
            sheet_count = 1
            worksheet = workbook.add_worksheet(f"Page {sheet_count}")
            worksheet.hide_gridlines(2)
            
            for col_idx, col_name in enumerate(column_names):
                worksheet.write(0, col_idx, col_name, header_format)
                
            row_in_sheet = 1
            MAX_ROWS_PER_SHEET = 1048500
            
            for idx, row in enumerate(rows):
                if task["abort_requested"]:
                    workbook.close()
                    raise Exception("Aborted by user")
                    
                if row_in_sheet >= MAX_ROWS_PER_SHEET:
                    sheet_count += 1
                    worksheet = workbook.add_worksheet(f"Page {sheet_count}")
                    worksheet.hide_gridlines(2)
                    for col_idx, col_name in enumerate(column_names):
                        worksheet.write(0, col_idx, col_name, header_format)
                    row_in_sheet = 1
                    
                for col_idx, val in enumerate(row):
                    if val is None:
                        worksheet.write_blank(row_in_sheet, col_idx, None, cell_format)
                    elif is_string_col[col_idx]:
                        worksheet.write_string(row_in_sheet, col_idx, str(val), text_format)
                    else:
                        if isinstance(val, (int, float)):
                            worksheet.write_number(row_in_sheet, col_idx, val, cell_format)
                        elif isinstance(val, bool):
                            worksheet.write_boolean(row_in_sheet, col_idx, val, cell_format)
                        else:
                            worksheet.write(row_in_sheet, col_idx, str(val), cell_format)
                            
                row_in_sheet += 1
                task["processed_rows"] += 1
                
            workbook.close()
            
        else:
            import csv
            delim = "\t" if format.startswith("tsv") else ","
            with open(file_path, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f, delimiter=delim)
                writer.writerow(column_names)
                for row in rows:
                    if task["abort_requested"]:
                        raise Exception("Aborted by user")
                    writer.writerow(row)
                    task["processed_rows"] += 1
                    
            if format.endswith(".gz"):
                gz_path = file_path + ".gz"
                with open(file_path, "rb") as f_in:
                    with gzip.open(gz_path, "wb") as f_out:
                        shutil.copyfileobj(f_in, f_out)
                os.remove(file_path)
                task["file_path"] = gz_path
                task["filename"] = final_filename + ".gz"
                
        task["status"] = "completed"
        task["elapsed"] = time.time() - task["start_time"]
        logger.info(f"[{task_id}] Export task finished successfully.")
    except Exception as e:
        logger.error(f"[{task_id}] Worker failed: {str(e)}")
        task["status"] = "failed"
        task["error"] = str(e)
        task["elapsed"] = time.time() - task["start_time"]
        if task_dir and os.path.exists(task_dir):
            try:
                shutil.rmtree(task_dir)
            except Exception:
                pass
    finally:
        if client:
            try:
                client.close()
            except Exception:
                pass

@app.post("/api/query/preview")
def preview_query(req: QueryRequest):
    logger.info("Executing query preview request...")
    raw_query = req.query.strip()
    if not raw_query:
        logger.warning("Empty query submitted")
        raise HTTPException(status_code=400, detail="Query cannot be empty")
        
    validate_select_query(raw_query)
    
    client = get_clickhouse_client()
    while raw_query.endswith(";"):
        raw_query = raw_query[:-1].strip()
    
    run_query = f"SELECT * FROM ({raw_query}) LIMIT 100"
    count_query = f"SELECT count() FROM ({raw_query})"
        
    try:
        # 1. Run preview
        logger.info(f"Running preview query: {run_query}")
        result = client.query(run_query)
        
        rows = []
        for row in result.result_rows:
            formatted_row = []
            for val in row:
                if val is None:
                    formatted_row.append(None)
                elif isinstance(val, (int, float, str, bool)):
                    formatted_row.append(val)
                else:
                    formatted_row.append(str(val))
            rows.append(formatted_row)
            
        # 2. Run optimized count
        total_rows = 0
        if count_query:
            logger.info(f"Running count query: {count_query}")
            count_res = client.query(count_query)
            total_rows = count_res.result_rows[0][0]
            logger.info(f"Optimized count result: {total_rows} total rows.")
            
        return {
            "columns": result.column_names,
            "types": [str(t) for t in result.column_types],
            "rows": rows,
            "total_rows": total_rows
        }
    except Exception as e:
        logger.error(f"Error executing preview/count: {str(e)}\n{traceback.format_exc()}")
        raise HTTPException(status_code=400, detail=f"Database error: {str(e)}")
    finally:
        try:
            client.close()
            logger.info("Connection closed for preview.")
        except Exception:
            pass

@app.post("/api/query/export")
def export_query(req: QueryRequest, format: str = "xlsx", background_tasks: BackgroundTasks = BackgroundTasks()):
    raw_query = req.query.strip()
    logger.info(f"Start exporting dataset. Format: {format}, Target Filename: {req.filename}")
    
    if not raw_query:
        logger.warning("Empty query submitted for export")
        raise HTTPException(status_code=400, detail="Query cannot be empty")
        
    validate_select_query(raw_query)
    cleanup_old_files()
    
    while raw_query.endswith(";"):
        raw_query = raw_query[:-1].strip()
        
    format = format.lower()
    valid_formats = ["xlsx", "csv", "tsv", "csv.gz", "tsv.gz"]
    if format not in valid_formats:
        logger.warning(f"Unsupported export format requested: {format}")
        raise HTTPException(status_code=400, detail=f"Unsupported format. Choose from: {valid_formats}")
    
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    base_filename = "".join(c for c in req.filename if c.isalnum() or c in (' ', '_', '-')).strip()
    base_filename = base_filename.replace(' ', '_')
    if not base_filename:
        base_filename = "export"
    
    # Generate task ID
    task_id = f"task_{timestamp}_{os.urandom(4).hex()}"
    EXPORT_TASKS[task_id] = {
        "status": "processing",
        "processed_rows": 0,
        "total_rows": 0,
        "start_time": time.time(),
        "elapsed": 0.0,
        "error": "",
        "file_path": "",
        "filename": "",
        "folder": "",
        "abort_requested": False
    }
    
    from threading import Thread
    thread = Thread(target=async_export_worker, args=(task_id, raw_query, base_filename, format, timestamp))
    thread.daemon = True
    thread.start()
    
    return {"task_id": task_id}

if __name__ == "__main__":
    import uvicorn
    if os.path.exists(TEMP_DIR):
        shutil.rmtree(TEMP_DIR)
    os.makedirs(TEMP_DIR, exist_ok=True)
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
