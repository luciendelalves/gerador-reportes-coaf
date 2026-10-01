"""
extrair_lote_coaf.py
====================

Extrai automaticamente os dados estruturados de TODOS os reportes já
escritos no arquivo .docx mestre (o arquivo único onde você acumula
todos os casos, um após o outro, no formato "NN) Nome completo: ...").

Gera um CSV pronto para ser importado pelo importador de lote do
gerador_coaf.py (aba "Importar lote existente de uma planilha/CSV").

------------------------------------------------------------------
COMO RODAR (na sua máquina, dentro da pasta do projeto)
------------------------------------------------------------------

1) Instalar a dependência (uma vez só):
   py -m pip install docx2txt --break-system-packages

2) Ajustar as duas variáveis abaixo (ARQUIVO_MESTRE e SAIDA_CSV) com
   os caminhos reais na sua máquina.

3) Rodar:
   py extrair_lote_coaf.py

4) Revisar o CSV gerado ANTES de importar no app — em especial as
   colunas "antecedente" e "reporte_anterior", que são inferidas por
   palavras-chave no texto (heurística), não são garantidas 100%
   corretas. Confira caso a caso antes de confiar cegamente nelas.

Você pode rodar este script de novo sempre que quiser atualizar o CSV
com os casos mais recentes — ele sempre processa o arquivo mestre
inteiro do zero, então sempre reflete todos os casos que estiverem
nele até aquele momento.
"""

import re
import os
import csv
import calendar

import docx2txt

# ==========================================================
# AJUSTE ESTES DOIS CAMINHOS ANTES DE RODAR
# ==========================================================

# Caminho do arquivo .docx mestre onde você acumula todos os reportes
ARQUIVO_MESTRE = "REPORTES AO COAF - MES DE AAAA.docx"

# Onde o CSV de saída deve ser salvo
SAIDA_CSV = "lote_extraido.csv"


# ==========================================================
# FUNÇÕES DE EXTRAÇÃO
# ==========================================================

# Cidade/UF da sede da comunicante — usada como fallback quando o local do
# apostador não pôde ser confirmado (ex: Netrin indisponível na consulta).
# Base: campo AgMun/AgUF do SISCOAF representa o "local do fato", que para
# operações inteiramente digitais corresponde ao local do estabelecimento
# que realizou o atendimento virtual — não o endereço do próprio apostador.
CIDADE_FALLBACK = os.environ.get("COAF_CIDADE_FALLBACK", "CIDADE")
UF_FALLBACK = "PE"

# Textos que indicam que a cidade não foi identificada no reporte original
MARCADORES_CIDADE_INDETERMINADA = {
    "INDETERMINADO", "INDETERMINADA", "NÃO IDENTIFICADA", "NAO IDENTIFICADA",
    "INDISPONÍVEL", "INDISPONIVEL", "VERIFICAR", "",
}


_CARACTERES_INICIO_FORMULA = ("=", "+", "-", "@", "\t", "\r")

def neutralizar_formula(texto):
    """Evita CSV/Formula Injection (OWASP): se o texto começar com um
    caractere que o Excel interpreta como início de fórmula, prefixa com
    apóstrofo para forçar leitura como texto literal ao abrir o CSV."""
    s = "" if texto is None else str(texto)
    if s and s[0] in _CARACTERES_INICIO_FORMULA:
        return "'" + s
    return s


def sanitizar_texto_siscoaf(texto):
    """Remove/substitui caracteres que o SISCOAF rejeita no campo
    Det/Informações Adicionais (erro 525: 'não pode conter os caracteres
    <, >, &, ", ', |'). Confirmado na prática com o retorno real de
    validação do SISCOAF em 29/07/2026 (casos com aspas em '"laranja"' e
    apóstrofo em 'D'Oeste' foram rejeitados)."""
    if not texto:
        return texto
    texto = texto.replace('&', ' e ')
    texto = texto.replace('"', '')
    texto = texto.replace("'", ' ')  # espaço no lugar do apóstrofo (D'Oeste -> D Oeste)
    texto = texto.replace('<', '').replace('>', '').replace('|', '')
    texto = re.sub(r' {2,}', ' ', texto)  # normaliza espaços duplicados que sobrarem
    return texto.strip()


PADRAO_CRIME_FINANCEIRO = (
    r'corrup[çc][ãa]o (ativa|passiva)|lavagem de dinheiro|crime financeiro|'
    r'crime contra o sistema financeiro|fraude (fiscal|financeira|contra|'
    r'[àa] licita[çc][ãa]o)|peculato|concuss[ãa]o|estelionato|'
    r'improbidade administrativa|sonega[çc][ãa]o fiscal|evas[ãa]o de divisas|'
    r'apropria[çc][ãa]o indébita|gest[ãa]o (fraudulenta|temer[áa]ria)'
)
PADRAO_CONTEXTO_ANTECEDENTE = (
    r'processo|indiciad[oa]|condenad[oa]|respons?de (como r[ée]u|a processo|a a[çc][ãa]o)|'
    r'r[ée]u em|investigad[oa]|inqu[ée]rito|execu[çc][ãa]o pen|a[çc][ãa]o (penal|civil p[úu]blica)'
)

def detectar_crime_financeiro(bloco):
    """Detecta antecedente de crime financeiro, exigindo que a palavra-chave
    apareça PERTO de um contexto de processo/antecedente real. Isso evita
    o falso positivo em que a frase padrão de fechamento de TODO reporte
    ('indícios de lavagem de dinheiro...') seria capturada como se fosse
    um antecedente criminal do apostador, quando na verdade é só o tipo de
    suspeita sendo comunicada agora."""
    for m in re.finditer(PADRAO_CRIME_FINANCEIRO, bloco, re.IGNORECASE):
        inicio_janela = max(0, m.start() - 200)
        fim_janela = min(len(bloco), m.end() + 200)
        janela = bloco[inicio_janela:fim_janela]
        if re.search(PADRAO_CONTEXTO_ANTECEDENTE, janela, re.IGNORECASE):
            return True
    return False


def extrair_campo(padrao, texto, default=""):
    """Busca o primeiro grupo capturado por um regex no texto do bloco."""
    m = re.search(padrao, texto)
    return m.group(1).strip() if m else default


def parse_valor_br(valor_str):
    """Converte '7.215.908,05' (formato BR) para inteiro (sem centavos),
    já que o leiaute do SISCOAF exige valores inteiros nos campos
    VlCred/VlDeb."""
    if not valor_str:
        return 0
    s = valor_str.replace(".", "").replace(",", ".")
    try:
        return int(round(float(s)))
    except ValueError:
        return 0


def periodo_para_datas(periodo_txt):
    """Converte 'MÊS 06/2026' em (DtInicio, DtFim) no formato dd/mm/aaaa,
    usando o primeiro e o último dia do mês informado."""
    m = re.search(r'(\d{1,2})/(\d{4})', periodo_txt)
    if not m:
        return "", ""
    mes, ano = int(m.group(1)), int(m.group(2))
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    return f"01/{mes:02d}/{ano}", f"{ultimo_dia:02d}/{mes:02d}/{ano}"


def processar_arquivo(caminho_docx):
    """Lê o .docx mestre, separa em blocos (um por caso, identificados
    pelo padrão 'NN) Nome completo:') e extrai os campos de cada um."""
    texto = docx2txt.process(caminho_docx)
    texto = texto.replace("\r\n", "\n")

    # Localiza o início de cada caso pelo padrão "NN) Nome completo:"
    inicios = [m.start() for m in re.finditer(r'\d{1,3}\)\s*Nome completo:', texto)]
    inicios.append(len(texto))

    casos = []
    for i in range(len(inicios) - 1):
        bloco = texto[inicios[i]:inicios[i + 1]].strip()
        if not bloco:
            continue

        numero_caso = extrair_campo(r'^(\d{1,3})\)', bloco)
        nome = extrair_campo(r'Nome completo:\s*(.+)', bloco)
        cpf = extrair_campo(r'CPF:\s*([\d\.\-]+)', bloco)
        data_nasc = extrair_campo(r'Data de nascimento:\s*([\d/]+)', bloco)
        cidade_uf_bruto = extrair_campo(r'Cidade/UF:\s*(.+)', bloco)

        # Prioriza a barra "/" como separador (padrão correto). Só usa
        # hífen "-" como alternativa quando NÃO houver barra nenhuma no
        # texto — isso evita quebrar cidades cujo próprio nome já tem
        # hífen (ex: "XANGRI-LÁ/RS" seria dividido errado em 3 partes se
        # dividíssemos sempre por ambos os caracteres ao mesmo tempo).
        if '/' in cidade_uf_bruto:
            partes_cidade = cidade_uf_bruto.split('/')
        else:
            partes_cidade = re.split(r'-', cidade_uf_bruto)
        cidade_extraida = partes_cidade[0].strip().upper() if partes_cidade else ""
        uf_extraida = partes_cidade[1].strip().upper() if len(partes_cidade) > 1 else ""

        # Fallback: se a cidade não foi identificada no reporte original
        # (Netrin indisponível na consulta, por exemplo), usa a sede da
        # Comunicante como "local do fato" — ver definição de AgMun/AgUF no
        # Manual de Integração do SISCOAF.
        if cidade_extraida in MARCADORES_CIDADE_INDETERMINADA or not uf_extraida:
            if cidade_extraida in MARCADORES_CIDADE_INDETERMINADA:
                cidade_extraida = CIDADE_FALLBACK
                uf_extraida = UF_FALLBACK
            # Caso tenha cidade real mas UF ausente (ex: "BETIM" sem UF
            # nenhuma no texto original), mantém a cidade e usa só a UF
            # fallback é arriscado demais para inferir — sinaliza para
            # revisão manual deixando a UF vazia mesmo, em vez de assumir.

        cidade_uf = f"{cidade_extraida}/{uf_extraida}" if uf_extraida else cidade_extraida
        pep = extrair_campo(r'PEP:\s*(SIM|NÃO|NAO)', bloco, "NÃO").upper().replace("NAO", "NÃO")
        periodo = extrair_campo(r'Período da\(s\) operação\(ões\):\s*(.+)', bloco)

        total_dep_str = extrair_campo(r'totalizaram o montante de R\$\s*([\d\.,]+)', bloco)
        qtd_dep_str = extrair_campo(r'registrados\s+([\d\.]+)\s+dep[óo]sitos', bloco)
        total_saq_str = extrair_campo(r'saques[^.]*?que somaram R\$\s*([\d\.,]+)', bloco)

        dt_inicio, dt_fim = periodo_para_datas(periodo)

        # Heurísticas por palavra-chave — REVISAR MANUALMENTE antes de confiar
        reporte_anterior = "SIM" if re.search(
            r'conduta reiterada|per[íi]odo(s)? consecutivo|reporte anterior|'
            r'segunda vez|segundo per[íi]odo|terceiro reporte|per[íi]odos analisados|'
            r'quarto per[íi]odo|per[íi]odos consecutivos',
            bloco, re.IGNORECASE) else "NÃO"

        antecedente = "SIM" if re.search(
            r'antecedente[s]? criminal|respons?de a processo|averiguado em inqu[ée]rito|'
            r'ação penal|execução pen',
            bloco, re.IGNORECASE) else "NÃO"

        # Detecção específica de antecedente ligado a CRIME FINANCEIRO
        # (corrupção, fraude, lavagem em processo à parte) — distinto do
        # antecedente genérico acima. Aplica o CodEnq 1410 no gerador_coaf.py
        # quando SIM. Usa janela de contexto (ver detectar_crime_financeiro)
        # para não confundir a frase padrão de fechamento ("indícios de
        # lavagem de dinheiro...", presente em TODO reporte) com um
        # antecedente criminal real do apostador.
        antecedente_crime_financeiro = "SIM" if detectar_crime_financeiro(bloco) else "NÃO"

        # Fracionamento: mesmo limiar usado no gerador principal (qtd_dep > 100).
        # Antes desta correção, essa informação nunca era extraída do texto,
        # então o fluxo de importação em massa nunca aplicava o CodEnq 1420
        # mesmo em casos com centenas de depósitos.
        try:
            qtd_dep_extraida = int(qtd_dep_str.replace(".", "")) if qtd_dep_str else 0
        except ValueError:
            qtd_dep_extraida = 0
        # Fracionamento: SIM se qtd_dep > 100 (limiar numérico objetivo) OU
        # se a palavra "fracionamento" aparecer explicitamente no texto —
        # cobre casos de fracionamento de SAQUES (não só depósitos) e
        # casos de julgamento qualitativo do analista mesmo com menos de
        # 100 depósitos. O Art. 25-XI (CodEnq 1420) cobre tanto "aporte"
        # quanto "retirada" de valores, então o mesmo código se aplica.
        tem_palavra_fracionamento = bool(re.search(r'fracionamento', bloco, re.IGNORECASE))
        fracionamento_detectado = "SIM" if (qtd_dep_extraida > 100 or tem_palavra_fracionamento) else "NÃO"

        # PEP relacionado: cobre o cenário em que o TITULAR não é PEP
        # direto (campo PEP: NÃO no cabeçalho), mas o texto menciona um
        # familiar ou sócio classificado como PEP (pai, cônjuge, sócio
        # etc.). A Resolução COAF nº 29/2017, Art. 2º, estende
        # expressamente o dever de atenção reforçada a "familiares" e
        # "estreitos colaboradores" de PEP, não só ao próprio PEP — por
        # isso aplicamos o mesmo CodEnq 1426 nesse cenário também.
        pep_relacionado_detectado = bool(re.search(r'PEP relacionad[oa]', bloco, re.IGNORECASE))

        casos.append({
            "numero_caso": numero_caso,
            "cpf": re.sub(r'\D', '', cpf),
            "nome": neutralizar_formula(sanitizar_texto_siscoaf(nome)),
            "data_nascimento": data_nasc,
            "cidade_uf": neutralizar_formula(sanitizar_texto_siscoaf(cidade_uf)),
            "pep": pep,
            "pep_relacionado": "SIM" if pep_relacionado_detectado else "NÃO",
            "periodo": periodo,
            "dt_inicio": dt_inicio,
            "dt_fim": dt_fim,
            "vl_cred": parse_valor_br(total_dep_str),
            "vl_deb": parse_valor_br(total_saq_str),
            "incompatibilidade": "SIM",  # presente em todos os casos do template atual
            "antecedente": antecedente,
            "antecedente_crime_financeiro": antecedente_crime_financeiro,
            "fracionamento": fracionamento_detectado,
            "qtd_dep": qtd_dep_extraida,
            "reporte_anterior": reporte_anterior,
            # O Excel só reconhece quebra de linha DENTRO de uma célula (mesmo
            # campo CSV) se ela estiver no padrão \r\n. Com \n puro, o Excel
            # espalha cada quebra em uma nova LINHA da planilha, quebrando o
            # "uma linha = um caso". Por isso convertemos aqui antes de gravar.
            # Também sanitizamos caracteres proibidos pelo SISCOAF (erro 525)
            # ANTES de aplicar a troca de quebra de linha.
            "det": neutralizar_formula(sanitizar_texto_siscoaf(bloco)).replace("\n", "\r\n"),
        })

    return casos


def main():
    if not os.path.isfile(ARQUIVO_MESTRE):
        print(f"ERRO: arquivo não encontrado: {ARQUIVO_MESTRE}")
        print("Ajuste a variável ARQUIVO_MESTRE no topo do script.")
        return

    print(f"Processando: {ARQUIVO_MESTRE}")
    casos = processar_arquivo(ARQUIVO_MESTRE)
    print(f"-> {len(casos)} caso(s) encontrado(s)")

    colunas = [
        "numero_caso", "cpf", "nome", "data_nascimento", "cidade_uf", "pep",
        "pep_relacionado", "periodo", "dt_inicio", "dt_fim", "vl_cred", "vl_deb",
        "incompatibilidade", "antecedente", "antecedente_crime_financeiro",
        "fracionamento", "qtd_dep", "reporte_anterior", "det",
    ]

    with open(SAIDA_CSV, "w", newline="", encoding="utf-8-sig") as f:
        # delimiter=";" porque o Excel em português (Brasil) usa vírgula como
        # separador decimal, e por padrão espera ";" como separador de coluna
        # em arquivos CSV abertos diretamente (duplo clique). Isso também
        # ajuda o Excel a reconhecer corretamente as quebras de linha internas
        # do campo 'det' como parte da mesma célula, em vez de linhas novas.
        writer = csv.DictWriter(f, fieldnames=colunas, delimiter=";")
        writer.writeheader()
        writer.writerows(casos)

    print(f"\nConcluído: {len(casos)} casos exportados para:\n{SAIDA_CSV}")
    print("\n⚠️  IMPORTANTE: revise o CSV antes de importar no app, especialmente")
    print("   as colunas 'antecedente' e 'reporte_anterior' (inferidas por texto,")
    print("   não são garantidas 100% precisas) e confirme os valores de")
    print("   vl_cred/vl_deb contra o reporte original em alguns casos de amostra.")


if __name__ == "__main__":
    main()
