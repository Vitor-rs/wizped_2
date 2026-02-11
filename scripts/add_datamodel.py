import xlwings as xw
import os
import sys

# Constants
FILE_NAME = "ficha_freq_gerador.xlsx"
TABLES = [
    "Tbl_Alunos",
    "Tbl_Agenda",
    "Tbl_Livros",
    "Tbl_Experiencia",
    "Tbl_Modalidades",
    "Tbl_Status",
    "Tbl_Contrato",
    "Tbl_Funcionarios",
    "Tbl_Horarios",
]


def add_data_model():
    file_path = os.path.abspath(FILE_NAME)
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return

    print(f"Opening {FILE_NAME}...")
    app = None
    wb = None
    try:
        # Launch invisible Excel
        app = xw.App(visible=False)
        wb = app.books.open(file_path)

        for table_name in TABLES:
            print(f"Processing {table_name}...")

            # Check if table exists (ListObject)
            found_table = False
            for ws in wb.sheets:
                try:
                    # xlwings collection access can be tricky, check safely
                    if table_name in [tbl.name for tbl in ws.tables]:
                        found_table = True
                        break
                except:
                    continue

            if not found_table:
                print(
                    f"  Warning: Table {table_name} not found in workbook sheets. Skipping."
                )
                continue

            # Check if connection already exists
            conn_name = f"WorksheetConnection_{table_name}"
            existing_conns = [c.Name for c in wb.api.Connections]

            if conn_name in existing_conns:
                print(f"  Connection {conn_name} already exists. Skipping.")
                continue

            print(f"  Adding {table_name} to Data Model...")

            # Construct connection string
            # "WORKSHEET;" indicates internal table
            conn_string = f"WORKSHEET;"

            try:
                wb.api.Connections.Add2(
                    conn_name,
                    "",
                    conn_string,
                    table_name,
                    7,  # xlCmdTable
                    True,  # CreateModelConnection
                    False,  # ImportRelationships
                )
                print("  Done.")
            except Exception as e:
                print(f"  Error adding connection: {e}")

        print("Saving...")
        wb.save()
        print("Success! All tables processed.")

    except Exception as e:
        print(f"Critical Error: {e}")
    finally:
        try:
            if wb:
                wb.close()
            if app:
                app.quit()
        except:
            pass


if __name__ == "__main__":
    add_data_model()
