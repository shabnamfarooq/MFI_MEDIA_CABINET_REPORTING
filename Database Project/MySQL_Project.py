import sqlite3

# Function to connect to the database
def connect_to_database(database_name):
    conn = sqlite3.connect(database_name)
    return conn

# Function to execute a query and fetch results
def execute_query(conn, query):
    cursor = conn.cursor()
    cursor.execute(query)
    rows = cursor.fetchall()
    return rows

# Function to print the count of rows in each table
def print_table_counts(conn):
    tables = ["Smartphones", "Brands"]
    for table in tables:
        query = f"SELECT COUNT(*) FROM {table};"
        result = execute_query(conn, query)
        print(f"Number of rows in {table}: {result[0][0]}")

# Function to print a sample of 3 rows from each table
def print_sample_rows(conn):
    tables = ["Smartphones", "Brands"]
    for table in tables:
        query = f"SELECT * FROM {table} LIMIT 3;"
        result = execute_query(conn, query)
        print(f"Sample rows from {table}:")
        for row in result:
            print(row)
        print()

# Function to perform a join query
def perform_join(conn):
    query = """
    SELECT s.*, b.Brand_Name
    FROM Smartphones AS s
    INNER JOIN Brands AS b ON Brand_ID = b.Brand_ID;
    """
    result = execute_query(conn, query)
    print("Join result:")
    for row in result:
        print(row)

# Main function
def main():
    database_name = "C:\\Users\\Home Easy\\Desktop\\My_database"
    conn = connect_to_database(database_name)

    # Print count of rows in each table
    print_table_counts(conn)
    print()

    # Print sample of 3 rows from each table
    print_sample_rows(conn)
    print()

    # Perform join
    perform_join(conn)

    # Close connection
    conn.close()

if __name__ == "__main__":
    main()
