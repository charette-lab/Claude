"""Connect to Azure SQL Server via ODBC Driver 18.

Prereq: install the Microsoft ODBC Driver 18 for SQL Server
    https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server

Python deps:
    pip install pyodbc

Credentials are read from environment variables so they are not committed:
    SQL_SERVER    e.g. msqlserver-aip.database.windows.net
    SQL_DATABASE  e.g. TruthDB_Copy
    SQL_USER      e.g. azureuser
    SQL_PASSWORD  the account password
"""

import os
import sys

import pyodbc


def build_connection_string() -> str:
    server = os.environ["SQL_SERVER"]
    database = os.environ["SQL_DATABASE"]
    user = os.environ["SQL_USER"]
    password = os.environ["SQL_PASSWORD"]
    return (
        "Driver={ODBC Driver 18 for SQL Server};"
        f"Server=tcp:{server},1433;"
        f"Database={database};"
        f"Uid={user};"
        f"Pwd={password};"
        "Encrypt=yes;"
        "TrustServerCertificate=no;"
        "Connection Timeout=30;"
    )


def main() -> int:
    conn_str = build_connection_string()
    with pyodbc.connect(conn_str) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT @@VERSION;")
        row = cursor.fetchone()
        print(row[0] if row else "no result")

        cursor.execute(
            "SELECT TABLE_SCHEMA, TABLE_NAME "
            "FROM INFORMATION_SCHEMA.TABLES "
            "WHERE TABLE_TYPE = 'BASE TABLE' "
            "ORDER BY TABLE_SCHEMA, TABLE_NAME;"
        )
        print("\nTables:")
        for schema, name in cursor.fetchall():
            print(f"  {schema}.{name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
