import pyodbc

access_path = r"C:\Users\USER\Documents\Python\LICIEL_Dossiers.mdb"
conn_str = f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={access_path};"

try:
    with pyodbc.connect(conn_str) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT TOP 5 * FROM Donnees_Dossiers")  # exemple de table
        for row in cursor.fetchall():
            print(row)
except Exception as e:
    print("❌ Erreur :", e)
