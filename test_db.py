import sqlite3
conn = sqlite3.connect('iseprep.db')
cursor = conn.cursor()
cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
tables = cursor.fetchall()
print("Tables:", tables)
for table in tables:
    cursor.execute(f'SELECT COUNT(*) FROM "{table[0]}"')
    count = cursor.fetchone()[0]
    print(f"Row count in {table[0]}: {count}")
    cursor.execute(f'SELECT * FROM "{table[0]}" LIMIT 1')
    print(f"Sample from {table[0]}:", cursor.fetchone())
conn.close()