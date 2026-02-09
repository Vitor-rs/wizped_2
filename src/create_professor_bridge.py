import xlwings as xw
import pandas as pd
import openpyxl
import os
import time

FILE_NAME = "ficha_freq_gerador.xlsx"
BRIDGE_SHEET = "BD_Vinculo_Professor"
BRIDGE_TABLE = "Tbl_Vinculo_Professor"


def migrate_professors():
    file_path = os.path.abspath(FILE_NAME)
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return

    # PART 1: Read Data (Skip if table already full to save time, or overwrite? Overwrite is safer)
    # ... (Reading part was successful, let's keep it but maybe optimize)
    # Actually, if table exists and has data, we might just want to fix relationships.
    # But let's assume we want to ensure data is there.

    print("Reading data with openpyxl...")
    try:
        wb_data = openpyxl.load_workbook(file_path, data_only=True)
        if "BD_Alunos" not in wb_data.sheetnames:
            print("BD_Alunos not found.")
            return

        ws_data = wb_data["BD_Alunos"]
        data = list(ws_data.values)
        wb_data.close()  # Close immediately

        if not data:
            print("No data in BD_Alunos")
            return

        header = data[0]
        rows = data[1:]

        df = pd.DataFrame(rows, columns=[str(h).strip() for h in header])

        col_id = next((c for c in df.columns if "ID_Aluno" in c), None)
        col_prof = next((c for c in df.columns if "Prof" in c), None)

        if not col_id or not col_prof:
            print("Columns not found.")
            df_migration = pd.DataFrame(columns=["ID_Aluno", "ID_Professor"])
        else:
            df_migration = df[[col_id, col_prof]].dropna()
            df_migration.columns = ["ID_Aluno", "ID_Professor"]
            df_migration = df_migration[
                pd.to_numeric(df_migration["ID_Aluno"], errors="coerce").notnull()
            ]
            df_migration = df_migration[
                pd.to_numeric(df_migration["ID_Professor"], errors="coerce").notnull()
            ]
            df_migration = df_migration.astype(int)

        print(f"Extracted {len(df_migration)} relationships.")

    except Exception as e:
        print(f"Error reading data: {e}")
        return

    # PART 2: Modify Workbook
    print(f"Opening {FILE_NAME} with xlwings...")
    app = xw.App(visible=False)
    try:
        wb = app.books.open(file_path)

        # Ensure Sheet/Table Exists with Data
        if BRIDGE_SHEET not in [s.name for s in wb.sheets]:
            ws_bridge = wb.sheets.add(BRIDGE_SHEET, after=wb.sheets[len(wb.sheets) - 1])
        else:
            ws_bridge = wb.sheets[BRIDGE_SHEET]

        # Write keys
        ws_bridge.range("A1").value = ["ID_Vinculo", "ID_Aluno", "ID_Professor"]
        if not df_migration.empty:
            df_migration.insert(0, "ID_Vinculo", range(1, len(df_migration) + 1))
            ws_bridge.range("A2").value = df_migration.values

        # Ensure ListObject
        last_row = max(2, len(df_migration) + 1)
        rng = ws_bridge.range(f"A1:C{last_row}")

        exists_tbl = False
        for tbl in ws_bridge.api.ListObjects:
            if tbl.Name == BRIDGE_TABLE:
                exists_tbl = True
                # Resize if needed? xlwings Resize not straight fwd on api
                tbl.Resize(rng.api)
                break

        if not exists_tbl:
            tbl = ws_bridge.api.ListObjects.Add(1, rng.api, 0, 1)
            tbl.Name = BRIDGE_TABLE
            tbl.TableStyle = "TableStyleMedium2"

        # Connection
        conn_name = f"WorksheetConnection_{BRIDGE_TABLE}"
        exists_conn = False
        for c in wb.api.Connections:
            if c.Name == conn_name:
                exists_conn = True

        if not exists_conn:
            wb.api.Connections.Add2(
                conn_name, "", "WORKSHEET;", BRIDGE_TABLE, 7, True, False
            )
            print(
                "Connection added. Creating relationships might fail if model isn't refreshed."
            )
            # We might need to save and reopen or force refresh?
            wb.api.Model.Refresh()
        else:
            print("Connection exists.")
            # Force refresh to ensure table is in model schema
            try:
                wb.api.Model.Refresh()
            except:
                pass

        # Relationships
        try:
            model = wb.api.Model
            rels = model.ModelRelationships
            tables = model.ModelTables

            # Debug: List tables in model
            # print("Model Tables:", [tables.Item(i).Name for i in range(1, tables.Count+1)])

            # Remove old
            try:
                # Find by iterating
                for k in range(1, rels.Count + 1):
                    r = rels.Item(k)
                    if (
                        r.ForeignKeyTable.Name == "Tbl_Alunos"
                        and r.ForeignKeyColumn.Name == "ID_Professor"
                    ):
                        r.Delete()
                        print("Deleted old relationship.")
                        break
            except:
                pass

            # Add New
            # Need to match names exactly.
            t_bridge = tables.Item(BRIDGE_TABLE)
            t_alunos = tables.Item("Tbl_Alunos")
            t_profs = tables.Item("Tbl_Professores")

            # Link 1: Bridge[ID_Aluno] -> Alunos[ID_Aluno (SponteWeb)]
            try:
                # Check existence
                exists = False
                for k in range(1, rels.Count + 1):
                    r = rels.Item(k)
                    if (
                        r.ForeignKeyTable.Name == BRIDGE_TABLE
                        and r.ForeignKeyColumn.Name == "ID_Aluno"
                    ):
                        exists = True

                if not exists:
                    c_fk = t_bridge.ModelTableColumns.Item("ID_Aluno")
                    c_pk = t_alunos.ModelTableColumns.Item("ID_Aluno (SponteWeb)")
                    rels.Add(c_fk, c_pk)
                    print("Linked Bridge->Alunos")
            except Exception as e:
                print(f"Link 1 error: {e}")

            # Link 2: Bridge[ID_Professor] -> Professores[ID_Professor]
            try:
                exists = False
                for k in range(1, rels.Count + 1):
                    r = rels.Item(k)
                    if (
                        r.ForeignKeyTable.Name == BRIDGE_TABLE
                        and r.ForeignKeyColumn.Name == "ID_Professor"
                    ):
                        exists = True

                if not exists:
                    c_fk = t_bridge.ModelTableColumns.Item("ID_Professor")
                    c_pk = t_profs.ModelTableColumns.Item("ID_Professor")
                    rels.Add(c_fk, c_pk)
                    print("Linked Bridge->Professores")
            except Exception as e:
                print(f"Link 2 error: {e}")

        except Exception as e:
            print(f"Model operations error: {e}")

        wb.save()
        print("Done.")

    except Exception as e:
        print(f"Available error: {e}")
    finally:
        try:
            wb.close()
            app.quit()
        except:
            pass


if __name__ == "__main__":
    migrate_professors()
