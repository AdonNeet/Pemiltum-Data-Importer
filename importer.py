import mysql.connector
import pandas as pd
from text import TextColor as tc

HOST = "localhost"
USER = "root"
PASSWORD = ""
DATABASE = "web_pemilu_fosti"
EXCEL_FILE = "data.xlsx"

def connect_mysql(host, user, password, database):
    db = mysql.connector.connect(
        host=host,
        user=user,
        password=password,
        database=database
    )
    return db

def insert_participant(email):
    cursor = db.cursor()
    sql = "INSERT INTO emails (email, created_at, updated_at) VALUES (%s, NOW(), NOW())"
    cursor.execute(sql, (email, ))
    db.commit()
    cursor.close()

if __name__ == "__main__":
    try:
        db = connect_mysql(
            host=HOST,
            user=USER,
            password=PASSWORD,
            database=DATABASE
        )
        print(f"{tc.GREEN}Success: Connected to database{tc.END}")
    except:
        print(f"{tc.RED}Error: Can't connect to database{tc.END}")
        exit(1)

    df = pd.read_excel(EXCEL_FILE)
    participants = tuple(df[['name', 'student_id', 'email']].values)
    for name, student_id, email in participants:
        try:
            student_id = student_id.upper()
            name = name.title()
            email = email.replace(" ", "")
            print(f"{tc.YELLOW}Action: Inserting {student_id} {name}{tc.END}")
            insert_participant(email)
            print(f"{tc.GREEN}Success: Inserting {student_id} {name}{tc.END}")
        except:
            print(f"{tc.RED}Error: Inserting {student_id} {name}{tc.END}")
            continue
    
    db.close()