import xlwings as xw
import os
import time

FILE_NAME = "ficha_freq_gerador.xlsx"


def link_monthly_sheet():
    file_path = os.path.abspath(FILE_NAME)
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return

    print(f"Opening {FILE_NAME}...")
    app = xw.App(visible=False)

    try:
        wb = app.books.open(file_path)
        print("Refreshing Model to ensure table schema is loaded...")
        try:
            wb.api.Model.Refresh()
            time.sleep(2)
        except:
            pass

        model = wb.api.Model
        rels = model.ModelRelationships
        tables = model.ModelTables

        # Link: Tbl_FichaMensal[ID_Aluno] -> Tbl_Alunos[ID_Aluno (SponteWeb)]
        t_many_name = "Tbl_FichaMensal"
        c_many_name = "ID_Aluno"
        t_one_name = "Tbl_Alunos"
        c_one_name = "ID_Aluno (SponteWeb)"

        # Check if exists
        exists = False
        for k in range(1, rels.Count + 1):
            r = rels.Item(k)
            try:
                if (
                    r.ForeignKeyTable.Name == t_many_name
                    and r.ForeignKeyColumn.Name == c_many_name
                    and r.PrimaryKeyTable.Name == t_one_name
                    and r.PrimaryKeyColumn.Name == c_one_name
                ):
                    exists = True
            except:
                pass

        if exists:
            print("Relationship already exists.")
        else:
            try:
                t_many = tables.Item(t_many_name)
                t_one = tables.Item(t_one_name)

                c_fk = t_many.ModelTableColumns.Item(c_many_name)
                c_pk = t_one.ModelTableColumns.Item(c_one_name)

                rels.Add(c_fk, c_pk)
                print(f"Created relationship: {t_many_name}->{t_one_name}")
                wb.save()
                print("Workbook saved.")
            except Exception as e:
                print(f"Error creating relationship: {e}")
                print(
                    "Tip: If error is generic, make sure 'ID_Aluno' exists in FichaMensal and types match."
                )

    except Exception as e:
        print(f"Critical Error: {e}")
    finally:
        try:
            wb.close()
            app.quit()
        except:
            pass


if __name__ == "__main__":
    link_monthly_sheet()
