# Python script to access SQLite database.

import sqlite3

# Enable for verbose debug logging (disabled by default)
g_EnableDebugMsg = True

'''
Function to connect to SQLite database and return a connection object.
'''
def connect_db(dbName):
    if not dbName:
        print('Error: Empty database name while connect')
        return None

    debug_msg("Connecting to SQLite database:%s" % dbName)

    try:
        # Connect to the SQLite database
        conn = sqlite3.connect(dbName)
        if not conn:
            print("Error: Failed to connect to SQLite database:%s" % dbName)
            return None

        # Create table if it does not exist yet
        isTableExists = exec_sql(conn, "SELECT name FROM sqlite_master WHERE type='table' and name='lastrun'")
        if not isTableExists:
            debug_msg("Setting up SQLite database:%s' ..." % dbName)
            exec_sql(conn, "CREATE TABLE lastrun (lasttime int, lastsubmit text)")
            exec_sql(conn, "INSERT INTO lastrun VALUES (0, '')")

        # Success
        return conn
    except sqlite3.Error as error:
        print("Error: Exception caught while conencting to SQLite database, error:'%s'" % str(error))
        return None

    # Failed
    return None

'''
Function to dis-connect from SQLite database.
'''
def disconnect_db(conn):
    if conn:
        debug_msg("Disconnecting SQLite database")
        conn.close()

'''
Generic function to execute any SQL query.
'''
def exec_sql(conn, sql):
    # Validate connection
    if not conn:
        print("Error: Connection is 'None' while excuting sql query:'%s'" % sql)
        return None

    # Validate connection
    if not sql:
        print("Error: sql is empty while executing query.")
        return None

    rows = None
    rowCount = None

    # Execute
    try:
        debug_msg("Executing SQL:'%s'" % sql)

        # Create a new cursor
        cur = conn.cursor()
        # Execute the sql statement
        cur.execute(sql)

        if "SELECT" in sql:
            rows = cur.fetchall() # get rows for SELECT query
        elif "UPDATE" in sql:
            rowCount = cur.rowcount # get rowCount for UPDATE query

        # Commit the changes to the database
        conn.commit()
        # Close communication with the database
        cur.close()

        # Success
        if rows:
            return rows
        elif rowCount is not None:
            return rowCount
    except sqlite3.Error as error:
        print("Error: Exception caught while executing sql query '%s', error:'%s'" % (sql, str(error)))
        conn.rollback() # rollback transaction
        return None

    # Failed
    return None

'''
Prints verbose debug message.
'''
def debug_msg(msg):
    if g_EnableDebugMsg:
        print(msg)
