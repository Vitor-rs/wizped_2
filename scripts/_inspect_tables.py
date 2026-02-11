"""Inspect table XML after openpyxl save to see namespace state."""

import zipfile
import shutil
import openpyxl

path = "docs/ficha_freq_gerador.xlsm"
test_path = "docs/_test_save.xlsm"
shutil.copy2(path, test_path)

# Do minimal openpyxl work and save
wb = openpyxl.load_workbook(test_path, keep_vba=True)
ws = wb["BD_Professores"]
ws.title = "BD_Funcionarios"
ws.cell(1, 1).value = "ID_Funcionario"
wb.save(test_path)
print("Saved via openpyxl.\n")

# Now inspect what openpyxl left in table XMLs
z = zipfile.ZipFile(test_path, "r")
for name in ["xl/tables/table5.xml", "xl/tables/table13.xml"]:
    print(f"\n{'=' * 60}")
    print(f"  {name}")
    print(f"{'=' * 60}")
    content = z.read(name).decode("utf-8")
    # Show first 500 chars to see namespace declarations
    print(content[:800])
    print("...")
z.close()

import os

os.remove(test_path)
