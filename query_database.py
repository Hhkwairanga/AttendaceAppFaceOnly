import sqlite3
import sqlite3

def update_email():
    try:
        # Connect to the SQLite database
        conn = sqlite3.connect('enrollment_data.db')
        cursor = conn.cursor()

        # Update email in the corpers table
        cursor.execute("UPDATE corpers SET email = 'harunahk5575@gmail.com' WHERE name = 'Haruna Hamidu'")
        conn.commit()
        print("Email updated in corpers table.")

        # Close the database connection
        conn.close()
    except sqlite3.Error as e:
        print("Error updating email:", e)

def query_database():
    try:
        # Connect to the SQLite database
        conn = sqlite3.connect('enrollment_data.db')
        cursor = conn.cursor()

        # Query the interns table
        print("Interns Table:")
        cursor.execute("SELECT * FROM interns")
        interns_data = cursor.fetchall()
        for row in interns_data:
            print(row)

        # Query the corpers table
        print("\nCorpers Table:")
        cursor.execute("SELECT * FROM corpers")
        corpers_data = cursor.fetchall()
        for row in corpers_data:
            print(row)

        # Close the database connection
        conn.close()
    except sqlite3.Error as e:
        print("Error querying database:", e)

if __name__ == "__main__":
    update_email()
    query_database()
