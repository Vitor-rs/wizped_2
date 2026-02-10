import xlwings as xw
import time
import os

FILE_NAME = "ficha_freq_gerador.xlsx"
BRIDGE_TABLE = "Tbl_Vinculo_Professor"


def link_bridge_table():
    file_path = os.path.abspath(FILE_NAME)
    print(f"Opening {FILE_NAME}...")
    app = xw.App(visible=False)

    try:
        wb = app.books.open(file_path)

        # Force Refresh to ensure Tbl_Vinculo_Professor is in Model
        print("Refreshing Model...")
        try:
            wb.api.Model.Refresh()
            time.sleep(5)  # Give it time?
        except Exception as e:
            print(f"Refresh warning: {e}")

        model = wb.api.Model
        rels = model.ModelRelationships
        tables = model.ModelTables

        print("Model Tables available:")
        # check if our table is there
        found = False
        for i in range(1, tables.Count + 1):
            t = tables.Item(i)
            # print(f" - {t.Name}")
            if t.Name == BRIDGE_TABLE:
                found = True

        if not found:
            print(
                f"Table {BRIDGE_TABLE} NOT found in Model yet. Try opening Excel manually and refreshing."
            )
            return

        print(f"Table {BRIDGE_TABLE} found. Creating relationships...")

        # 1. Bridge -> Alunos
        try:
            t_bridge = tables.Item(BRIDGE_TABLE)
            t_alunos = tables.Item("Tbl_Alunos")

            c_fk = t_bridge.ModelTableColumns.Item("ID_Aluno")
            c_pk = t_alunos.ModelTableColumns.Item("ID_Aluno (SponteWeb)")

            rels.Add(c_fk, c_pk)
            print("Linked Bridge->Alunos")
        except Exception as e:
            print(f"Link Bridge->Alunos failed (might exist): {e}")

        # 2. Bridge -> Professores
        try:
            t_bridge = tables.Item(BRIDGE_TABLE)
            t_profs = tables.Item("Tbl_Professores")

            c_fk = t_bridge.ModelTableColumns.Item("ID_Professor")
            c_pk = t_profs.ModelTableColumns.Item("ID_Professor")

            rels.Add(c_fk, c_pk)
            print("Linked Bridge->Professores")
        except Exception as e:
            print(f"Link Bridge->Professores failed (might exist): {e}")

        # 3. Delete old direct link
        try:
            for k in range(1, rels.Count + 1):
                r = rels.Item(k)
                if (
                    r.PrimaryKeyTable.Name == "Tbl_Professores"
                    and r.ForeignKeyTable.Name == "Tbl_Alunos"
                ):
                    r.Delete()
                    print("Deleted old direct link Alunos->Professores")
        except:
            pass

        wb.save()
        print("Done.")

    except Exception as e:
        print(f"Error: {e}")
    finally:
        try:
            wb.close()
            app.quit()
        except:
            pass


if __name__ == "__main__":
    link_bridge_table()
