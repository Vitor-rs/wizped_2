import pandas as pd
import os

FILE_NAME = "ficha_freq_gerador.xlsx"


def analyze_data():
    file_path = os.path.abspath(FILE_NAME)
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return

    print(f"Reading {FILE_NAME}...")
    try:
        # Read all tables (sheets are named same as tables in this clean file)
        xls = pd.read_excel(file_path, sheet_name=None, engine="openpyxl")

        # 1. Analyze Tbl_Horarios
        if "BD_Horarios" in xls:
            df_horarios = xls["BD_Horarios"]
            print("\n=== Analysis: Tbl_Horarios (BD_Horarios) ===")
            print(f"Rows: {len(df_horarios)}")
            print("Columns:", list(df_horarios.columns))

            # Check uniqueness of Hora
            if "Hora" in df_horarios.columns:
                unique_horas = df_horarios["Hora"].nunique()
                print(f"Unique 'Hora' values: {unique_horas}")
                if unique_horas == len(df_horarios):
                    print("SUCCESS: 'Hora' column is UNIQUE. Can be used as a Key.")
                else:
                    print(
                        "WARNING: 'Hora' column contains DUPLICATES. Cannot be a One-side Key."
                    )
                    print(df_horarios["Hora"].value_counts())
            else:
                print("ERROR: Column 'Hora' not found.")

        # 2. Analyze Tbl_Agenda Compatibility
        if "BD_Agenda" in xls and "BD_Horarios" in xls:
            df_agenda = xls["BD_Agenda"]
            df_horarios = xls["BD_Horarios"]

            print("\n=== Analysis: Agenda -> Horarios (Relationship) ===")
            if "Hora" in df_agenda.columns and "Hora" in df_horarios.columns:
                # Check data types
                print(f"Agenda[Hora] Type: {df_agenda['Hora'].dtype}")
                print(f"Horarios[Hora] Type: {df_horarios['Hora'].dtype}")

                # Check Referential Integrity (Orphans)
                # Convert to string for comparison if needed, or keeping explicit
                agenda_times = set(df_agenda["Hora"].dropna())
                master_times = set(df_horarios["Hora"].dropna())

                orphans = agenda_times - master_times
                if orphans:
                    print(
                        f"WARNING: Found {len(orphans)} times in Agenda NOT present in Horarios Master List."
                    )
                    print(f"Orphan examples: {list(orphans)[:5]}")
                else:
                    print("SUCCESS: All times in Agenda exist in Horarios Master List.")
            else:
                print("Skipping check: 'Hora' column missing in one of the tables.")

        # 3. Quick stats on others
        print("\n=== Quick Stats ===")
        for sheet, df in xls.items():
            if sheet not in ["BD_Horarios", "BD_Agenda"]:
                print(f"{sheet}: {len(df)} rows")

    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    analyze_data()
