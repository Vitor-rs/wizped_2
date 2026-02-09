import xlwings as xw
import os
import time

FILE_NAME = "ficha_freq_gerador.xlsx"

# 1. Tables to Add
TABLES = [
    "Tbl_Alunos",
    "Tbl_Agenda",
    "Tbl_Livros",
    "Tbl_Experiencia",
    "Tbl_Modalidades",
    "Tbl_Status",
    "Tbl_Contrato",
    "Tbl_Professores",
    "Tbl_Horarios",
    "Tbl_Vinculo_Professor",
    "Tbl_FichaMensal",
]

# 2. Relationships (Many -> One)
# (ManyTable, ManyCol, OneTable, OneCol)
RELATIONSHIPS = [
    ("Tbl_Alunos", "ID_Livro", "Tbl_Livros", "ID_Livro"),
    ("Tbl_Alunos", "ID_Status", "Tbl_Status", "ID_Status"),
    ("Tbl_Alunos", "ID_Contrato", "Tbl_Contrato", "ID_Contrato"),
    ("Tbl_Alunos", "ID_Experiencia", "Tbl_Experiencia", "ID_Experiencia"),
    ("Tbl_Alunos", "ID_Modalidade", "Tbl_Modalidades", "ID_Modalidade"),
    # ('Tbl_Alunos', 'ID_Professor', 'Tbl_Professores', 'ID_Professor'), # REMOVED (Replaced by Bridge)
    ("Tbl_Agenda", "ID_Aluno (SponteWeb)", "Tbl_Alunos", "ID_Aluno (SponteWeb)"),
    ("Tbl_Agenda", "Hora", "Tbl_Horarios", "Hora"),
    # Bridge Relationships
    ("Tbl_Vinculo_Professor", "ID_Aluno", "Tbl_Alunos", "ID_Aluno (SponteWeb)"),
    ("Tbl_Vinculo_Professor", "ID_Professor", "Tbl_Professores", "ID_Professor"),
]


def rebuild_model():
    file_path = os.path.abspath(FILE_NAME)
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return

    print(f"Opening {FILE_NAME}...")
    app = xw.App(visible=False)

    try:
        wb = app.books.open(file_path)

        # --- PHASE 1: Add Tables ---
        print("\n--- Phase 1: Adding Tables ---")
        for tbl_name in TABLES:
            conn_name = f"WorksheetConnection_{tbl_name}"

            exists = False
            for c in wb.api.Connections:
                if c.Name == conn_name:
                    exists = True

            if exists:
                print(f"  {tbl_name} already connected.")
            else:
                try:
                    wb.api.Connections.Add2(
                        conn_name, "", "WORKSHEET;", tbl_name, 7, True, False
                    )
                    print(f"  {tbl_name} added.")
                except Exception as e:
                    print(f"  Failed adding {tbl_name}: {e}")

        # Refresh
        print("Refreshing Model...")
        try:
            wb.api.Model.Refresh()
            time.sleep(2)
        except:
            print("Refresh soft fail.")

        # --- PHASE 2: Relationships Retry Loop ---
        print("\n--- Phase 2: Relationships (Retry Loop) ---")
        model = wb.api.Model
        rels = model.ModelRelationships
        model_tables = model.ModelTables

        def get_col(t_name, c_name):
            try:
                return model_tables.Item(t_name).ModelTableColumns.Item(c_name)
            except:
                return None

        def rel_exists(t1, c1, t2, c2):
            for k in range(1, rels.Count + 1):
                r = rels.Item(k)
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

        max_attempts = 3
        for attempt in range(max_attempts):
            print(f"Attempt {attempt + 1}/{max_attempts}...")
            all_done = True

            for t_many, c_many, t_one, c_one in RELATIONSHIPS:
                if rel_exists(t_many, c_many, t_one, c_one):
                    continue

                c_fk = get_col(t_many, c_many)
                c_pk = get_col(t_one, c_one)

                if c_fk and c_pk:
                    try:
                        rels.Add(c_fk, c_pk)
                        print(f"  Created: {t_many}->{t_one}")
                    except Exception as e:
                        print(f"  Error creating {t_many}->{t_one}: {e}")
                        all_done = False
                else:
                    print(f"  Waiting for columns: {t_many}->{t_one}")
                    all_done = False

            if all_done:
                print("All relationships verified.")
                break

            if attempt < max_attempts - 1:
                print("Waiting 5s for Model refresh...")
                time.sleep(5)
                try:
                    wb.api.Model.Refresh()
                except:
                    pass

        wb.save()
        print("\nRebuild Complete.")

    except Exception as e:
        print(f"Critical Error: {e}")
    finally:
        try:
            wb.close()
            app.quit()
        except:
            pass


if __name__ == "__main__":
    rebuild_model()
