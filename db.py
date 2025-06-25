import pyodbc
import configparser
from typing import Optional, Any, List, Tuple


class DatabaseConnection:
    def __init__(self, config: configparser.ConfigParser):
        self.conn: Optional[pyodbc.Connection] = None
        self.cursor: Optional[pyodbc.Cursor] = None
        self.config = config

    def connect(self) -> bool:
        try:
            drivers = pyodbc.drivers()
            print("Available drivers:", drivers)

            if "ODBC Driver 18 for SQL Server" in drivers:
                driver = "ODBC Driver 18 for SQL Server"
            elif "ODBC Driver 17 for SQL Server" in drivers:
                driver = "ODBC Driver 17 for SQL Server"
            else:
                raise Exception("No suitable SQL Server driver found")

            server = self.config.get("database", "host")
            if self.config.get("database", "port"):
                server = f"{server},{self.config.get("database", "port")}"

            conn_str = (
                f"Driver={driver};"
                f"Server={server};"
                f"Database={self.config.get("database", "database")};"
                f"UID={self.config.get("database", "user")};"
                f"PWD={self.config.get("database", "password")};"
                "TrustServerCertificate=yes;"
                "Connection Timeout=30;"
                "Login Timeout=30;"
            )

            print("Attempting to connect with connection string:")
            safe_conn_str = conn_str.replace(
                self.config.get("database", "password"), "****"
            )
            print(safe_conn_str)

            self.conn = pyodbc.connect(conn_str)
            self.cursor = self.conn.cursor()
            print("Successfully connected to the database!")
            return True

        except pyodbc.Error as e:
            print(f"Error connecting to the database: {str(e)}")
            print("\nCommon solutions:")
            print("1. Verify SQL Server is running and accessible")
            print("2. Check if SQL Server Authentication is enabled")
            print("3. Ensure port 1433 is open and not blocked by firewall")
            print("4. Verify SQL credentials are correct")
            print("5. Try connecting with Windows Authentication if available")
            return False
        except Exception as e:
            print(f"General error: {str(e)}")
            return False

    def _ensure_connection(self) -> None:
        """Ensure we have a valid connection and cursor"""
        if self.cursor is None or self.conn is None:
            raise Exception("No database connection. Call connect() first.")

        # Optional: Test if connection is still alive
        try:
            self.cursor.execute("SELECT 1")
            self.cursor.fetchone()
        except pyodbc.Error:
            raise Exception("Database connection lost. Please reconnect.")

    def execute_query(
        self, query: str, params: Optional[Tuple] = None
    ) -> List[pyodbc.Row]:
        print(query)
        """Execute a SELECT query and return all results"""
        self._ensure_connection()

        if self.cursor is None:  # Type guard for mypy
            raise Exception("Cursor is None")

        try:
            if params:
                self.cursor.execute(query, params)
            else:
                self.cursor.execute(query)
            return self.cursor.fetchall()
        except pyodbc.Error as e:
            print(f"Error executing query: {str(e)}")
            raise

    def execute_query_as_dict(
        self, query: str, params: Optional[Tuple] = None
    ) -> List[dict]:
        """Execute a SELECT query and return results as list of dictionaries"""
        self._ensure_connection()

        if self.cursor is None:  # Type guard for mypy
            raise Exception("Cursor is None")

        try:
            if params:
                self.cursor.execute(query, params)
            else:
                self.cursor.execute(query)

            # Get column names
            columns = (
                [column[0] for column in self.cursor.description]
                if self.cursor.description
                else []
            )

            # Convert rows to dictionaries
            rows = self.cursor.fetchall()
            return [dict(zip(columns, row)) for row in rows]
        except pyodbc.Error as e:
            print(f"Error executing query: {str(e)}")
            raise

    def execute_query_as_tuples(
        self, query: str, params: Optional[Tuple] = None
    ) -> List[Tuple[Any, ...]]:
        """Execute a SELECT query and return results as list of tuples (if you need tuple type)"""
        self._ensure_connection()

        if self.cursor is None:  # Type guard for mypy
            raise Exception("Cursor is None")

        try:
            if params:
                self.cursor.execute(query, params)
            else:
                self.cursor.execute(query)

            rows = self.cursor.fetchall()
            # Convert Row objects to tuples
            return [tuple(row) for row in rows]
        except pyodbc.Error as e:
            print(f"Error executing query: {str(e)}")
            raise

    def execute_scalar(self, query: str, params: Optional[Tuple] = None) -> Any:
        """Execute a query and return a single value"""
        self._ensure_connection()

        if self.cursor is None:  # Type guard for mypy
            raise Exception("Cursor is None")

        try:
            if params:
                self.cursor.execute(query, params)
            else:
                self.cursor.execute(query)
            result = self.cursor.fetchone()
            return result[0] if result else None
        except pyodbc.Error as e:
            print(f"Error executing scalar query: {str(e)}")
            raise

    def execute_non_query(self, query: str, params: Optional[Tuple] = None) -> int:
        """Execute INSERT/UPDATE/DELETE and return rows affected"""
        self._ensure_connection()

        if self.cursor is None:  # Type guard for mypy
            raise Exception("Cursor is None")

        try:
            if params:
                self.cursor.execute(query, params)
            else:
                self.cursor.execute(query)

            if self.conn is None:  # Type guard for mypy
                raise Exception("Connection is None")

            self.conn.commit()
            return self.cursor.rowcount
        except pyodbc.Error as e:
            if self.conn:
                self.conn.rollback()
            print(f"Error executing non-query: {str(e)}")
            raise

    def test_connection(self) -> None:
        """Test the database connection with comprehensive checks"""
        self._ensure_connection()

        if self.cursor is None:  # Type guard for mypy
            raise Exception("Cursor is None")

        try:
            # Test server connection
            server_name = self.execute_scalar("SELECT @@SERVERNAME")
            print(f"Connected to server: {server_name}")

            # Test database connection
            db_name = self.execute_scalar("SELECT DB_NAME()")
            print(f"Connected to database: {db_name}")

            # Get SQL Server version
            version = self.execute_scalar("SELECT @@VERSION")
            print(f"SQL Server version: {version}")

            # Test user permissions
            user_info = self.execute_query(
                """
                SELECT SUSER_SNAME() as username, 
                       USER_NAME() as database_user
            """
            )
            if user_info:
                print(f"Connected as: {user_info[0][0]}")
                print(f"Database user: {user_info[0][1]}")

        except Exception as e:
            print(f"Error executing test query: {str(e)}")
            raise

    def is_connected(self) -> bool:
        """Check if we have a valid connection"""
        if self.cursor is None or self.conn is None:
            return False

        try:
            self.cursor.execute("SELECT 1")
            self.cursor.fetchone()
            return True
        except pyodbc.Error:
            return False

    def close(self) -> None:
        """Close the database connection"""
        if self.cursor:
            self.cursor.close()
            self.cursor = None
        if self.conn:
            self.conn.close()
            self.conn = None
            print("Database connection closed.")

    def __enter__(self):
        """Context manager entry"""
        if not self.connect():
            raise Exception("Failed to connect to database")
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit"""
        self.close()
