import xlwings as xw
import os
import time

FILE_NAME = "ficha_freq_gerador.xlsx"


def create_history_tables():
    file_path = os.path.abspath(FILE_NAME)
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return

    print(f"Opening {FILE_NAME}...")
    app = xw.App(visible=False)

    try:
        wb = app.books.open(file_path)

        # --- 1. Tbl_TipoOcorrencia ---
        SHEET_TIPO = "BD_TipoOcorrencia"
        TABLE_TIPO = "Tbl_TipoOcorrencia"

        if SHEET_TIPO not in [s.name for s in wb.sheets]:
            ws_tipo = wb.sheets.add(SHEET_TIPO, after=wb.sheets[len(wb.sheets) - 1])
            # Headers
            ws_tipo.range("A1").value = ["ID_TipoOcorrencia", "Descricao"]
            # Default Data
            data_tipo = [
                [1, "Matrícula"],
                [2, "Rematrícula"],
                [3, "Entrega de Material"],
            ]
            ws_tipo.range("A2").value = data_tipo

            # Table
            last_row = len(data_tipo) + 1
            rng = ws_tipo.range(f"A1:B{last_row}")
            tbl = ws_tipo.api.ListObjects.Add(1, rng.api, 0, 1)
            tbl.Name = TABLE_TIPO
            tbl.TableStyle = "TableStyleMedium2"
            print(f"Table {TABLE_TIPO} created with defaults.")
        else:
            print(f"Table {TABLE_TIPO} already exists.")

        # --- 2. Tbl_Historico ---
        SHEET_HIST = "BD_Historico"
        TABLE_HIST = "Tbl_Historico"

        if SHEET_HIST not in [s.name for s in wb.sheets]:
            ws_hist = wb.sheets.add(SHEET_HIST, after=wb.sheets[len(wb.sheets) - 1])
            # Headers
            ws_hist.range("A1").value = [
                "ID_Historico",
                "ID_Aluno",
                "ID_Livro",
                "ID_TipoOcorrencia",
                "Data",
                "Observacao",
            ]

            # Table (Empty initially)
            # Add one empty row to stabilize ListObject
            ws_hist.range("A2").value = [None, None, None, None, None, None]

            rng = ws_hist.range("A1:F2")
            tbl = ws_hist.api.ListObjects.Add(1, rng.api, 0, 1)
            tbl.Name = TABLE_HIST
            tbl.TableStyle = "TableStyleMedium2"
            print(f"Table {TABLE_HIST} created.")
        else:
            print(f"Table {TABLE_HIST} already exists.")

        # --- 3. Data Model Connections ---
        print("Adding connections...")
        for t in [TABLE_TIPO, TABLE_HIST]:
            conn_name = f"WorksheetConnection_{t}"
            exists = False
            for c in wb.api.Connections:
                if c.Name == conn_name:
                    exists = True

            if not exists:
                wb.api.Connections.Add2(conn_name, "", "WORKSHEET;", t, 7, True, False)
                print(f"Connection for {t} added.")

        # Refresh
        print("Refreshing Model...")
        try:
            wb.api.Model.Refresh()
            time.sleep(2)
        except:
            pass

        # --- 4. Relationships ---
        # Historico -> Alunos
        # Historico -> Livros
        # Historico -> TipoOcorrencia

        RELATIONSHIPS = [
            (TABLE_HIST, "ID_Aluno", "Tbl_Alunos", "ID_Aluno (SponteWeb)"),
            (TABLE_HIST, "ID_Livro", "Tbl_Livros", "ID_Livro"),
            (TABLE_HIST, "ID_TipoOcorrencia", TABLE_TIPO, "ID_TipoOcorrencia"),
        ]

        model = wb.api.Model
        rels = model.ModelRelationships
        model_tables = model.ModelTables

        print("Creating relationships...")
        for t_many, c_many, t_one, c_one in RELATIONSHIPS:
            try:
                # Check exist
                exists = False
                for k in range(1, rels.Count + 1):
                    r = rels.Item(k)
                    if (
                        r.ForeignKeyTable.Name == t_many
                        and r.ForeignKeyColumn.Name == c_many
                    ):
                        exists = True

                if not exists:
                    t_m = model_tables.Item(t_many)
                    t_o = model_tables.Item(t_one)
                    c_fk = t_m.ModelTableColumns.Item(c_many)
                    c_pk = t_o.ModelTableColumns.Item(c_one)
                    rels.Add(c_fk, c_pk)
                    print(f"Created: {t_many}->{t_one}")
                else:
                    print(f"Exists: {t_many}->{t_one}")
            except Exception as e:
                print(f"Error linking {t_many}->{t_one}: {e}")

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
    create_history_tables()
