#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
rebuild_model.py — Wizped Office: Rebuild completo do Data Model

Adiciona TODAS as tabelas (ListObjects) ao Modelo de Dados do Excel
e cria TODOS os relacionamentos de forma idempotente.

USO:
    python scripts/rebuild_model.py "docs/ficha_freq_gerador.xlsm"

DEPENDÊNCIA:
    pip install xlwings
    (Requer Excel instalado — usa COM/API)
"""

import xlwings as xw
import os
import sys
import time

# ============================================================
# CONFIGURAÇÃO
# ============================================================

DEFAULT_FILE = os.path.join("docs", "ficha_freq_gerador.xlsm")

# Relacionamentos: (FK_Table, FK_Column, PK_Table, PK_Column)
RELATIONSHIPS = [
    # Alunos → Lookups
    ("Tbl_Alunos", "ID_Livro", "Tbl_Livros", "ID_Livro"),
    ("Tbl_Alunos", "ID_Status", "Tbl_Status", "ID_Status"),
    ("Tbl_Alunos", "ID_Contrato", "Tbl_Contrato", "ID_Contrato"),
    ("Tbl_Alunos", "ID_Experiencia", "Tbl_Experiencia", "ID_Experiencia"),
    ("Tbl_Alunos", "ID_Modalidade", "Tbl_Modalidades", "ID_Modalidade"),
    # Agenda → Alunos, Horarios
    ("Tbl_Agenda", "ID_Aluno (SponteWeb)", "Tbl_Alunos", "ID_Aluno (SponteWeb)"),
    ("Tbl_Agenda", "Hora", "Tbl_Horarios", "Hora"),
    # Bridge: Aluno ↔ Professor (N:N)
    ("Tbl_Vinculo_Professor", "ID_Aluno", "Tbl_Alunos", "ID_Aluno (SponteWeb)"),
    ("Tbl_Vinculo_Professor", "ID_Professor", "Tbl_Professores", "ID_Professor"),
    # Historico → Alunos, Livros, TipoOcorrencia
    ("Tbl_Historico", "ID_Aluno", "Tbl_Alunos", "ID_Aluno (SponteWeb)"),
    ("Tbl_Historico", "ID_Livro", "Tbl_Livros", "ID_Livro"),
    ("Tbl_Historico", "ID_TipoOcorrencia", "Tbl_TipoOcorrencia", "ID_TipoOcorrencia"),
    # FichaMensal → Alunos
    ("Tbl_FichaMensal", "ID_Aluno", "Tbl_Alunos", "ID_Aluno (SponteWeb)"),
    # Livros → Experiencia (default experience per book)
    ("Tbl_Livros", "ID_Experiencia_Padrao", "Tbl_Experiencia", "ID_Experiencia"),
]

MAX_ATTEMPTS = 3
WAIT_BETWEEN_ATTEMPTS = 5  # seconds


# ============================================================
# HELPERS
# ============================================================


def discover_tables(wb):
    """Auto-descobre todas as ListObjects (Excel Tables) no workbook."""
    tables = []
    for ws in wb.sheets:
        try:
            for tbl in ws.tables:
                tables.append(tbl.name)
        except Exception:
            continue
    return sorted(set(tables))


def get_existing_connections(wb):
    """Retorna set de nomes de conexões existentes."""
    names = set()
    try:
        for c in wb.api.Connections:
            names.add(c.Name)
    except Exception:
        pass
    return names


def add_table_to_model(wb, table_name, existing_conns):
    """Adiciona uma tabela ao Data Model via WorksheetConnection. Retorna True se adicionou."""
    conn_name = f"WorksheetConnection_{table_name}"

    if conn_name in existing_conns:
        print(f"  ─ {table_name} já conectada.")
        return False

    try:
        wb.api.Connections.Add2(
            conn_name,
            "",  # Description
            "WORKSHEET;",  # Connection string (internal table)
            table_name,  # Command text
            7,  # xlCmdTable
            True,  # CreateModelConnection
            False,  # ImportRelationships
        )
        print(f"  ✓ {table_name} adicionada ao modelo.")
        return True
    except Exception as e:
        print(f"  ✗ Erro ao adicionar {table_name}: {e}")
        return False


def refresh_model(wb, wait=3):
    """Refresh do Data Model com wait."""
    print(f"\n⟳ Refreshing Model (aguardando {wait}s)...")
    try:
        wb.api.Model.Refresh()
        time.sleep(wait)
        print("  ✓ Model refreshed.")
    except Exception as e:
        print(f"  ⚠ Refresh falhou (pode ser normal): {e}")


def rel_exists(rels, fk_table, fk_col, pk_table, pk_col):
    """Verifica se um relacionamento já existe."""
    try:
        for k in range(1, rels.Count + 1):
            r = rels.Item(k)
            try:
                if (
                    r.ForeignKeyTable.Name == fk_table
                    and r.ForeignKeyColumn.Name == fk_col
                    and r.PrimaryKeyTable.Name == pk_table
                    and r.PrimaryKeyColumn.Name == pk_col
                ):
                    return True
            except Exception:
                continue
    except Exception:
        pass
    return False


def get_model_column(model_tables, table_name, col_name):
    """Retorna ModelTableColumn ou None se não encontrar."""
    try:
        tbl = model_tables.Item(table_name)
        return tbl.ModelTableColumns.Item(col_name)
    except Exception:
        return None


def create_relationship(rels, model_tables, fk_table, fk_col, pk_table, pk_col):
    """Tenta criar um relacionamento. Retorna: 'created', 'exists', 'skipped', ou 'error'."""
    # Check existence
    if rel_exists(rels, fk_table, fk_col, pk_table, pk_col):
        return "exists"

    # Get columns
    col_fk = get_model_column(model_tables, fk_table, fk_col)
    col_pk = get_model_column(model_tables, pk_table, pk_col)

    if not col_fk or not col_pk:
        missing = []
        if not col_fk:
            missing.append(f"{fk_table}[{fk_col}]")
        if not col_pk:
            missing.append(f"{pk_table}[{pk_col}]")
        return f"skipped (coluna não encontrada: {', '.join(missing)})"

    try:
        rels.Add(col_fk, col_pk)
        return "created"
    except Exception as e:
        return f"error ({e})"


# ============================================================
# MAIN
# ============================================================


def main():
    # --- Resolve file path ---
    if len(sys.argv) > 1:
        file_path = os.path.abspath(sys.argv[1])
    else:
        file_path = os.path.abspath(DEFAULT_FILE)

    if not os.path.exists(file_path):
        print(f"✗ Arquivo não encontrado: {file_path}")
        sys.exit(1)

    print("=" * 60)
    print("WIZPED — Rebuild Data Model")
    print("=" * 60)
    print(f"Arquivo: {file_path}")
    print()

    app = None
    wb = None

    try:
        # --- Open ---
        print("Abrindo Excel...")
        app = xw.App(visible=False)
        wb = app.books.open(file_path)
        print("  ✓ Workbook aberto.\n")

        # ========================================
        # FASE 1: Descobrir e adicionar tabelas
        # ========================================
        print("=" * 60)
        print("FASE 1: Adicionando tabelas ao Data Model")
        print("=" * 60)

        discovered = discover_tables(wb)
        print(f"Tabelas encontradas ({len(discovered)}):")
        for t in discovered:
            print(f"  • {t}")
        print()

        existing_conns = get_existing_connections(wb)
        added_count = 0
        for table_name in discovered:
            if add_table_to_model(wb, table_name, existing_conns):
                added_count += 1

        if added_count > 0:
            print(f"\n  {added_count} tabela(s) adicionada(s).")
        else:
            print(f"\n  Todas as tabelas já estavam no modelo.")

        # ========================================
        # FASE 2: Refresh
        # ========================================
        refresh_model(wb, wait=3)

        # ========================================
        # FASE 3: Criar relacionamentos
        # ========================================
        print("\n" + "=" * 60)
        print("FASE 3: Criando relacionamentos")
        print("=" * 60)

        model = wb.api.Model
        rels = model.ModelRelationships
        model_tables = model.ModelTables

        # Show existing model tables
        try:
            model_table_names = []
            for i in range(1, model_tables.Count + 1):
                model_table_names.append(model_tables.Item(i).Name)
            print(f"Tabelas no modelo ({len(model_table_names)}):")
            for t in sorted(model_table_names):
                print(f"  • {t}")
            print()
        except Exception:
            print("  ⚠ Não foi possível listar tabelas do modelo.\n")

        # Retry loop
        for attempt in range(1, MAX_ATTEMPTS + 1):
            print(f"--- Tentativa {attempt}/{MAX_ATTEMPTS} ---")
            pending = []
            created_this_round = 0

            for fk_tbl, fk_col, pk_tbl, pk_col in RELATIONSHIPS:
                label = f"{fk_tbl}[{fk_col}] → {pk_tbl}[{pk_col}]"
                result = create_relationship(
                    rels, model_tables, fk_tbl, fk_col, pk_tbl, pk_col
                )

                if result == "created":
                    print(f"  ✓ {label}")
                    created_this_round += 1
                elif result == "exists":
                    print(f"  ─ {label} (já existe)")
                elif result.startswith("skipped"):
                    print(f"  ⚠ {label} — {result}")
                    # Only retry if it's a timing issue (table not yet in model)
                    if "coluna não encontrada" in result:
                        pending.append((fk_tbl, fk_col, pk_tbl, pk_col))
                else:
                    print(f"  ✗ {label} — {result}")
                    pending.append((fk_tbl, fk_col, pk_tbl, pk_col))

            if not pending:
                print(f"\n✓ Todos os relacionamentos verificados.")
                break

            if attempt < MAX_ATTEMPTS:
                print(
                    f"\n  {len(pending)} pendente(s). "
                    f"Aguardando {WAIT_BETWEEN_ATTEMPTS}s antes de retry..."
                )
                time.sleep(WAIT_BETWEEN_ATTEMPTS)
                refresh_model(wb, wait=2)
                # Re-acquire references after refresh
                model = wb.api.Model
                rels = model.ModelRelationships
                model_tables = model.ModelTables
            else:
                print(f"\n⚠ {len(pending)} relacionamento(s) não puderam ser criados:")
                for fk_tbl, fk_col, pk_tbl, pk_col in pending:
                    print(f"    • {fk_tbl}[{fk_col}] → {pk_tbl}[{pk_col}]")
                print(
                    "  Isso pode ser normal se a tabela/coluna ainda não existe no workbook."
                )

        # ========================================
        # FASE 4: Salvar
        # ========================================
        print(f"\nSalvando workbook...")
        wb.save()
        print("✓ Workbook salvo com sucesso.")

        # Summary
        print("\n" + "=" * 60)
        print("RESUMO")
        print("=" * 60)
        try:
            final_rels = wb.api.Model.ModelRelationships
            print(f"Tabelas no modelo: {model_tables.Count}")
            print(f"Relacionamentos:   {final_rels.Count}")

            print("\nRelacionamentos ativos:")
            for k in range(1, final_rels.Count + 1):
                r = final_rels.Item(k)
                try:
                    print(
                        f"  {r.ForeignKeyTable.Name}[{r.ForeignKeyColumn.Name}]"
                        f" → {r.PrimaryKeyTable.Name}[{r.PrimaryKeyColumn.Name}]"
                    )
                except Exception:
                    print(f"  (relacionamento {k}: erro ao ler)")
        except Exception as e:
            print(f"  ⚠ Erro ao gerar resumo: {e}")

        print("\n✓ Rebuild concluído!")

    except Exception as e:
        print(f"\n✗ Erro crítico: {e}")
        import traceback

        traceback.print_exc()
        sys.exit(1)

    finally:
        try:
            if wb:
                wb.close()
        except Exception:
            pass
        try:
            if app:
                app.quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
