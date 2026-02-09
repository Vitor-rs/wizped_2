import xlwings as xw
import os

FILE_NAME = "ficha_freq_gerador.xlsx"
OLD_NAME = "BP_FichaMensal"
NEW_NAME = "BD_FichaMensal"


def rename_sheet():
    file_path = os.path.abspath(FILE_NAME)
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return

    print(f"Opening {FILE_NAME}...")
    app = xw.App(visible=False)

    try:
        wb = app.books.open(file_path)

        found = False
        for sheet in wb.sheets:
            if sheet.name == OLD_NAME:
                sheet.name = NEW_NAME
                print(f"Renamed sheet '{OLD_NAME}' to '{NEW_NAME}'.")
                found = True
                break
            elif sheet.name == NEW_NAME:
                print(f"Sheet already named '{NEW_NAME}'.")
                found = True
                break

        if not found:
            print(f"Sheet '{OLD_NAME}' not found.")

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
    rename_sheet()
