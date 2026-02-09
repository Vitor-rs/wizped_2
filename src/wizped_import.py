#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
wizped_import.py — Parser de PDF do Sponte para Wizped Office

Extrai Nº Matrícula, Nome e Situação do relatório "Dados do Cadastro"
gerado pelo Sponte Web e salva em CSV para importação via VBA.

USO:
    python wizped_import.py "C:\caminho\relatorio_Cadastro_Alunos.pdf"

SAÍDA:
    Cria arquivo CSV no mesmo diretório do PDF:
    <nome_do_pdf>_parsed.csv

DEPENDÊNCIA:
    pip install pdfplumber
"""

import sys
import os
import csv
import unicodedata

def normalize(text):
    """Remove espaços extras e normaliza unicode"""
    if not text:
        return ""
    return ' '.join(str(text).strip().split())

def extract_students(pdf_path):
    """Extrai dados de alunos do PDF do Sponte"""
    try:
        import pdfplumber
    except ImportError:
        print("ERRO: pdfplumber não instalado. Execute: pip install pdfplumber")
        sys.exit(1)

    students = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if not row or not row[0]:
                        continue

                    matricula = normalize(row[0])
                    nome = normalize(row[1]) if len(row) > 1 and row[1] else ''
                    situacao = normalize(row[2]) if len(row) > 2 and row[2] else ''

                    # Pular headers
                    if 'Matrícula' in matricula:
                        continue

                    # Validar: matrícula deve ser numérica
                    try:
                        mat_int = int(matricula)
                        if mat_int > 0 and nome and situacao in ('Ativo', 'Desistente', 'Interessado', 'Trancado'):
                            students.append({
                                'id': mat_int,
                                'nome': nome,
                                'situacao': situacao
                            })
                    except (ValueError, TypeError):
                        continue

    return students

def main():
    if len(sys.argv) < 2:
        print("USO: python wizped_import.py <caminho_do_pdf>")
        sys.exit(1)

    pdf_path = sys.argv[1]

    if not os.path.exists(pdf_path):
        print(f"ERRO: Arquivo não encontrado: {pdf_path}")
        sys.exit(1)

    print(f"Processando: {pdf_path}")
    students = extract_students(pdf_path)
    print(f"Alunos extraidos: {len(students)}")

    # Salvar CSV
    csv_path = os.path.splitext(pdf_path)[0] + "_parsed.csv"
    with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f, delimiter=';')
        writer.writerow(['ID', 'Nome', 'Situacao'])
        for s in students:
            writer.writerow([s['id'], s['nome'], s['situacao']])

    print(f"CSV salvo: {csv_path}")
    print(f"OK:{len(students)}")  # Marcador para VBA parsear

if __name__ == '__main__':
    main()
