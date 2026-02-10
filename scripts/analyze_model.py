import xlwings as xw
import os

FILE_NAME = "ficha_freq_gerador.xlsx"


def analyze_model_structure():
    file_path = os.path.abspath(FILE_NAME)
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return

    print(f"Opening {FILE_NAME}...")
    app = None
    wb = None
    try:
        app = xw.App(visible=False)
        wb = app.books.open(file_path)

        try:
            model = wb.api.Model
        except:
            print("Error accessing Model.")
            return

        model_tables = model.ModelTables
        print(f"Identified {model_tables.Count} tables in Data Model:\n")

        for i in range(1, model_tables.Count + 1):
            tbl = model_tables.Item(i)
            print(f"### Table: {tbl.Name}")
            print("  Columns:")
            for j in range(1, tbl.ModelTableColumns.Count + 1):
                col = tbl.ModelTableColumns.Item(j)
                print(f"    - {col.Name}")
            print("")

    except Exception as e:
        print(f"Error: {e}")
    finally:
        try:
            if wb:
                wb.close()
            if app:
                app.quit()
        except:
            pass


if __name__ == "__main__":
    analyze_model_structure()
