# 📊 ClickHouse Query Executor & Excel Exporter 🚀

A fast, lightweight web interface built with **FastAPI**, **ClickHouse Connect**, and **XlsxWriter** to execute SQL queries on a ClickHouse database and export the results directly into beautiful, compressed Excel files (`.xlsx`).

---

## ✨ Features

- **⚡ Fast Query Execution:** Connects directly to ClickHouse and streams queries.
- **📁 Excel Export (`.xlsx`):** Formats query output into styled, ready-to-read Excel sheets.
- **🗜️ Automatic Compression:** Automatically packages exports in ZIP/GZIP formats for large datasets.
- **💻 Minimalist UI:** Simple and beautiful web interface to input queries, manage history, and download reports.
- **🛡️ Secure:** Uses `.env` configuration to keep database credentials hidden.

---

## 🛠️ Tech Stack

* [FastAPI](https://fastapi.tiangolo.com/) - Web framework
* [ClickHouse Connect](https://github.com/ClickHouse/clickhouse-connect) - Database driver
* [XlsxWriter](https://xlsxwriter.readthedocs.io/) - Excel creation
* [Python-dotenv](https://github.com/theofidry/django-dotenv) - Environment config
* HTML/CSS/Vanilla JS - Frontend

---

## ⚙️ Configuration

Create a `.env` file in the root directory to store your ClickHouse credentials:

```ini
CLICKHOUSE_HOST=localhost
CLICKHOUSE_PORT=8123
CLICKHOUSE_USER=default
CLICKHOUSE_PASSWORD=your_password
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
uvicorn main:app --reload
```

### 3. Access the Web UI
Open your browser and navigate to:
* 🖥️ **Web Interface:** `http://127.0.0.1:8000`
* 📖 **API Docs (Swagger):** `http://127.0.0.1:8000/docs`

---

## 📌 Usage Guide

1. Enter your **ClickHouse SQL Query** in the editor.
2. Provide a custom **Filename** for the export.
3. Click **Execute & Export** 🚀
4. Download the generated Excel file from your history panel!
