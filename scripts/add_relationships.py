import xlwings as xw
import os
import sys

# Constants
FILE_NAME = "ficha_freq_gerador.xlsx"

# Relationship Mapping
# (Many Side Table, Many Side Column) -> (One Side Table, One Side Column)
RELATIONSHIPS = [
    # Already done:
    # ('Tbl_Alunos', 'ID_Livro', 'Tbl_Livros', 'ID_Livro'),
    # ('Tbl_Alunos', 'ID_Status', 'Tbl_Status', 'ID_Status'),
    # ...
    # New final relationship
    # Tbl_Agenda[Hora] (Many) -> Tbl_Horarios[Hora] (One)
    ("Tbl_Agenda", "Hora", "Tbl_Horarios", "Hora")
]


def add_all_relationships():
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
        except Exception as e:
            print("Error accessing Workbook.Model.")
            return

        model_tables = model.ModelTables
        model_rels = model.ModelRelationships

        # Helper to find table/column
        def get_column(table_name, col_name):
            try:
                tbl = model_tables.Item(table_name)
                return tbl.ModelTableColumns.Item(col_name)
            except:
                return None

        # Helper to check if relationship exists
        def rel_exists(t1, c1, t2, c2):
            for k in range(1, model_rels.Count + 1):
                r = model_rels.Item(k)
                try:
                    if (
                        r.ForeignKeyTable.Name == t1
                        and r.ForeignKeyColumn.Name == c1
                        and r.PrimaryKeyTable.Name == t2
                        and r.PrimaryKeyColumn.Name == c2
                    ):
                        return True
                except:
                    pass
            return False

        print("Creating relationships...")

        for t_many, c_many, t_one, c_one in RELATIONSHIPS:
            print(f"Processing: {t_many}[{c_many}] -> {t_one}[{c_one}]")

            col_fk = get_column(t_many, c_many)
            col_pk = get_column(t_one, c_one)

            if not col_fk or not col_pk:
                print(
                    f"  Error: Could not find columns for {t_many}->{t_one}. Skipping."
                )
                continue

            if rel_exists(t_many, c_many, t_one, c_one):
                print("  Relationship already exists.")
            else:
                try:
                    model_rels.Add(col_fk, col_pk)
                    print("  Created successfully.")
                except Exception as e:
                    print(f"  Failed: {e}")

        print("Saving...")
        wb.save()
        print("Success! All relationships processed.")

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
    add_all_relationships()
