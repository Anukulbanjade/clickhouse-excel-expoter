# 📊 ClickHouse Query Executor & Excel Exporter 🚀

A fast, lightweight web interface built with **FastAPI**, **ClickHouse Connect**, and **XlsxWriter** to execute SQL queries on a ClickHouse database and export the results directly into beautiful, compressed Excel files (`.xlsx`).

---

## ✨ Features

- **⚡ Fast Query Execution:** Connects directly to ClickHouse and streams queries.
- **🔌 Connection Management:** Displays live connection states and provides explicit `Connect`, `Disconnect`, and `Reconnect` action toggles.
- **📁 Excel Export (`.xlsx`):** Formats query output into styled, ready-to-read Excel sheets.
- **🗜️ Automatic Compression:** Automatically packages exports in ZIP/GZIP formats for large datasets.
- **⚡ Optimized Workload:** Operates on-demand (loading exports and connection status updates only when tabs are changed or queries run) to eliminate idle background server load.
- **💻 Minimalist UI:** Simple and beautiful web interface to input queries, manage history, and download reports.

---

## 🛠️ Tech Stack

* [FastAPI](https://fastapi.tiangolo.com/) - Web framework
* [ClickHouse Connect](https://github.com/ClickHouse/clickhouse-connect) - Database driver
* [XlsxWriter](https://xlsxwriter.readthedocs.io/) - Excel creation
* [Cryptography](https://cryptography.io/en/latest/) - RSA-OAEP & Fernet ciphers
* [Python-dotenv](https://github.com/theofidry/django-dotenv) - Environment config
* HTML/CSS/Vanilla JS - Frontend

---

## ⚙️ Configuration

Create a `.env` file in the root directory (based on `.env_temp` template) to store your ClickHouse credentials and admin access key:

```ini
CLICKHOUSE_HOST=localhost
CLICKHOUSE_PORT=8123
CLICKHOUSE_USER=default
CLICKHOUSE_PASSWORD=your_password
ACCESS_KEY=admin
```

---

## 🚀 Getting Started

### 1. Install Dependencies
Make sure you have Python installed, then run:
```bash
pip install -r requirements.txt
```

### 2. Run the Application
Start the FastAPI server using Uvicorn:
```bash
python main.py
```

### 3. Access the Web UI
Open your browser and navigate to:
* 🖥️ **Web Interface:** `http://127.0.0.1:8000`
* 📖 **API Docs (Swagger):** `http://127.0.0.1:8000/docs`

---

## 📌 Usage Guide

1. Authenticate using the **Access Key** configured in your `.env` file.
2. Verify that ClickHouse shows `Connected` in the header bar.
3. Enter your **ClickHouse SQL Query** in the editor.
4. Provide a custom **Filename** for the export.
5. Click **Run Preview & Count** to preview rows and verify database connectivity.
6. Export as **Excel (XLSX)**, **CSV**, or **TSV** and download your query cache files from the **Recent Exports** panel!
