"""
verificar_gaps.py
==================

Script auxiliar: lê o CSV já gerado pelo extrair_lote_coaf.py e mostra
quais números de caso (numero_caso) NÃO aparecem nele, comparando com a
sequência esperada (1 até o maior número encontrado).

Isso ajuda a descobrir quais casos o extrator não conseguiu capturar do
documento mestre (geralmente por pequena variação de formatação nesse
bloco específico, como "087)" em vez de "87)", espaço extra, etc.

COMO RODAR:
   py verificar_gaps.py
"""

import csv

# Ajuste se o CSV estiver em outro lugar
CAMINHO_CSV = "lote_extraido.csv"


def main():
    numeros_encontrados = []
    with open(CAMINHO_CSV, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, delimiter=";")
        for linha in reader:
            try:
                numeros_encontrados.append(int(linha["numero_caso"]))
            except (ValueError, KeyError):
                pass

    if not numeros_encontrados:
        print("Nenhum número de caso foi lido do CSV. Confira o caminho do arquivo.")
        return

    numeros_encontrados_set = set(numeros_encontrados)
    maior_numero = max(numeros_encontrados_set)

    print(f"Total de casos no CSV: {len(numeros_encontrados)}")
    print(f"Números únicos encontrados: {len(numeros_encontrados_set)}")
    print(f"Maior número de caso encontrado: {maior_numero}")
    print()

    faltando = [n for n in range(1, maior_numero + 1) if n not in numeros_encontrados_set]
    if faltando:
        print(f"⚠️  {len(faltando)} número(s) da sequência 1-{maior_numero} NÃO aparecem no CSV:")
        print(faltando)
        print()
        print("Provável causa: o script de extração não reconheceu o início desses")
        print("casos no documento mestre (formatação ligeiramente diferente do padrão")
        print("'NN) Nome completo:'). Abra o documento mestre, procure esses números")
        print("de caso, e confira se o cabeçalho está escrito exatamente como nos outros.")
    else:
        print("✅ Nenhum número faltando na sequência — todos os casos de 1 até "
              f"{maior_numero} foram capturados.")

    # Verificar duplicados também, por precaução
    duplicados = [n for n in numeros_encontrados_set if numeros_encontrados.count(n) > 1]
    if duplicados:
        print(f"\n⚠️  Números que aparecem MAIS DE UMA VEZ no CSV: {duplicados}")


if __name__ == "__main__":
    main()
