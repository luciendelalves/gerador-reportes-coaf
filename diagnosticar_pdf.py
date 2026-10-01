"""
Diagnostico de extracao de PDF - Gerador COAF.

Uso:
    py diagnosticar_pdf.py "caminho\\do\\arquivo.pdf"

Nao imprime CPFs nem nomes - so estrutura (versoes, tabelas encontradas,
cabecalhos, quantidade de digitos por CPF lido) para comparar entre maquinas.
"""
import sys
import re

print("=== Diagnostico de extracao de PDF - Gerador COAF ===")

try:
    import pdfplumber
    print(f"pdfplumber: {pdfplumber.__version__}")
except Exception as e:
    print(f"ERRO ao importar pdfplumber: {e}")
    sys.exit(1)

try:
    import pdfminer
    print(f"pdfminer.six: {pdfminer.__version__}")
except Exception as e:
    print(f"(nao consegui checar versao do pdfminer.six: {e})")

if len(sys.argv) < 2:
    print('\nUso: py diagnosticar_pdf.py "caminho\\do\\arquivo.pdf"')
    sys.exit(0)

caminho = sys.argv[1]

with pdfplumber.open(caminho) as pdf:
    print(f"\nPaginas: {len(pdf.pages)}")
    for i, page in enumerate(pdf.pages):
        tables_linhas = page.extract_tables({
            "vertical_strategy": "lines",
            "horizontal_strategy": "lines",
            "text_x_tolerance": 1.5,
        })
        tables_texto = page.extract_tables({"text_x_tolerance": 1.5})
        print(f"\n-- Pagina {i + 1} --")
        print(f"Tabelas (estrategia linhas): {len(tables_linhas)}")
        print(f"Tabelas (estrategia texto): {len(tables_texto)}")

        tabelas = tables_linhas if tables_linhas else tables_texto
        if not tabelas:
            print("  Nenhuma tabela encontrada por nenhuma estrategia nesta pagina.")

        for t_idx, table in enumerate(tabelas):
            if not table or len(table) < 2:
                continue
            header_idx = None
            for r_idx, row in enumerate(table):
                if row and any(str(c).upper().strip() in ['CPF', 'NOME', 'NOME COMPLETO'] for c in row if c):
                    header_idx = r_idx
                    break
            if header_idx is None:
                print(f"  Tabela {t_idx}: {len(table)} linhas, cabecalho CPF/NOME NAO encontrado.")
                print(f"  Primeira linha (so para ver formato, confira se nao tem dado pessoal antes de colar aqui): {table[0]}")
            else:
                headers = [str(h).strip().upper() if h else '' for h in table[header_idx]]
                print(f"  Tabela {t_idx}: cabecalho encontrado na linha {header_idx}: {headers}")
                print(f"  Total de linhas de dados: {len(table) - header_idx - 1}")
                cpf_idx = next((idx for idx, h in enumerate(headers) if 'CPF' in h), None)
                if cpf_idx is None:
                    print("  Coluna CPF nao encontrada no cabecalho.")
                else:
                    tamanhos = []
                    for row in table[header_idx + 1:]:
                        if row and cpf_idx < len(row) and row[cpf_idx]:
                            digs = re.sub(r'\D', '', str(row[cpf_idx]))
                            tamanhos.append(len(digs))
                    print(f"  Quantidade de digitos por CPF lido nas primeiras linhas (deve ser 11): {tamanhos[:15]}")

print("\n=== Fim do diagnostico ===")
