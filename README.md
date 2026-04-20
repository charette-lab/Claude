# SQL Server Connection

Minimal script that connects to Azure SQL Server using the ODBC Driver 18.

## 1. Install the ODBC driver

Follow the Microsoft guide for your OS:
<https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server>

## 2. Install Python deps

```bash
pip install -r requirements.txt
```

## 3. Provide credentials

Copy `.env.example` to `.env` and fill in real values, or export the variables
directly in your shell:

```bash
export SQL_SERVER=msqlserver-aip.database.windows.net
export SQL_DATABASE=TruthDB_Copy
export SQL_USER=azureuser
export SQL_PASSWORD='...'
```

## 4. Run

```bash
python connect.py
```

The script prints `@@VERSION` and lists the base tables in the database.

## Security note

Do not commit credentials. `.env` is gitignored. If a password has been
pasted into chat or a ticket, rotate it in the Azure portal.
