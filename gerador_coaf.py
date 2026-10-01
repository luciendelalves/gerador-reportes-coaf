import streamlit as st
import pdfplumber
import re
import json
import os
import random
import subprocess
from io import BytesIO
from datetime import date, datetime
import calendar
import xml.etree.ElementTree as ET

try:
    import gspread
    from google.oauth2.credentials import Credentials as UserCredentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request as GoogleAuthRequest
    GSPREAD_DISPONIVEL = True
except ImportError:
    GSPREAD_DISPONIVEL = False

try:
    import win32crypt  # DPAPI — criptografia dos dados locais amarrada à conta Windows
    DPAPI_DISPONIVEL = True
except ImportError:
    DPAPI_DISPONIVEL = False

try:
    from docx import Document
    from docx.shared import Pt
    DOCX_DISPONIVEL = True
except ImportError:
    DOCX_DISPONIVEL = False


st.set_page_config(page_title="Gerador de Reportes COAF", page_icon="📋", layout="wide")
st.title("📋 Gerador de Reportes COAF")
st.markdown("---")

# COAF_DATA_DIR permite redirecionar onde os 3 arquivos de dados locais
# são lidos/gravados — usado pelos testes automatizados (test_gerador_coaf.py)
# para nunca tocar os arquivos reais do projeto. Em uso normal (sem essa
# variável de ambiente definida), o comportamento é o de sempre: os 3
# arquivos ficam na mesma pasta do gerador_coaf.py.
_DIRETORIO_DADOS = os.environ.get("COAF_DATA_DIR") or os.path.dirname(__file__)
HISTORICO_FILE = os.path.join(_DIRETORIO_DADOS, "historico_casos.json")
LOTE_CASOS_FILE = os.path.join(_DIRETORIO_DADOS, "lote_casos.json")
# Arquivo PERMANENTE com o texto completo de todo caso já adicionado ao
# lote — ao contrário de lote_casos.json (fila do ciclo atual, esvaziada a
# cada envio), este nunca é apagado automaticamente. É o que substitui o
# documento mestre externo como registro consultável dos reportes.
ARQUIVO_CASOS_FILE = os.path.join(_DIRETORIO_DADOS, "arquivo_casos.json")

# ==========================================================
# SEGURANÇA: proteção de arquivos locais sensíveis
# ==========================================================
# historico_casos.json, lote_casos.json e token_google.json contêm dados
# pessoais de apostadores (LGPD) e/ou credenciais de acesso ao Google
# Sheets. Além do .gitignore (que só protege contra commit no Git),
# aplicamos nesta máquina e em qualquer outra onde este arquivo for
# copiado:
#   1) Criptografia via DPAPI do Windows (CryptProtectData) — o conteúdo
#      só pode ser lido pela MESMA conta Windows que gravou o arquivo
#      (não depende de senha extra para gerenciar, nem é portável entre
#      contas/máquinas — por isso nunca copie esses 3 arquivos de uma
#      máquina para outra esperando que funcionem lá).
#   2) ACL do NTFS restringindo o arquivo ao usuário atual (best-effort).
# Arquivos antigos em texto puro são lidos normalmente e migram sozinhos
# para o formato protegido no próximo salvamento — não precisa apagar
# nada manualmente. Se o pywin32 não estiver instalado na máquina, os
# arquivos continuam funcionando em texto puro (comportamento anterior),
# com um aviso visível na tela.

_DPAPI_UI_FORBIDDEN = 0x1  # CRYPTPROTECT_UI_FORBIDDEN — nunca abre popup do Windows

def _dpapi_proteger(dados_bytes, descricao):
    return win32crypt.CryptProtectData(dados_bytes, descricao, None, None, None, _DPAPI_UI_FORBIDDEN)

def _dpapi_desproteger(dados_bytes):
    _descricao, dados = win32crypt.CryptUnprotectData(dados_bytes, None, None, None, _DPAPI_UI_FORBIDDEN)
    return dados

def _restringir_acesso_arquivo(caminho):
    """Restringe o arquivo ao usuário Windows atual (best-effort: nunca
    interrompe o app se falhar, ex.: fora do Windows ou sem permissão)."""
    if os.name != "nt" or not os.path.exists(caminho):
        return
    usuario = os.environ.get("USERNAME", "")
    if not usuario:
        return
    try:
        subprocess.run(
            ["icacls", caminho, "/inheritance:r", "/grant:r", f"{usuario}:F"],
            capture_output=True, timeout=5, check=False,
        )
    except Exception:
        pass

def _ler_json_protegido(caminho, padrao, descricao="coaf_dados"):
    """Lê um JSON local que pode estar em texto puro (arquivos antigos) ou
    protegido por DPAPI (arquivos novos) — migra transparentemente para o
    formato protegido no próximo salvamento."""
    if not os.path.exists(caminho):
        return padrao
    with open(caminho, "rb") as f:
        bruto = f.read()
    if not bruto:
        return padrao
    try:
        return json.loads(bruto.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError):
        pass
    if not DPAPI_DISPONIVEL:
        raise RuntimeError(
            f"O arquivo {caminho} está criptografado (DPAPI), mas o módulo pywin32 "
            "não está instalado nesta máquina. Rode: py -m pip install pywin32"
        )
    try:
        dados = _dpapi_desproteger(bruto)
    except Exception as e:
        raise RuntimeError(
            f"Não foi possível descriptografar {caminho}: {e}. Esse arquivo só pode "
            "ser lido pela mesma conta Windows que o criou — se ele foi copiado de "
            "outra máquina/usuário, restaure o backup local ou recomece do zero."
        )
    return json.loads(dados.decode("utf-8"))

def _salvar_json_protegido(caminho, obj, descricao="coaf_dados"):
    dados = json.dumps(obj, ensure_ascii=False, indent=2).encode("utf-8")
    if DPAPI_DISPONIVEL:
        dados = _dpapi_proteger(dados, descricao)
    with open(caminho, "wb") as f:
        f.write(dados)
    _restringir_acesso_arquivo(caminho)

_CARACTERES_INICIO_FORMULA = ("=", "+", "-", "@", "\t", "\r")

def neutralizar_formula(valor):
    """Evita CSV/Formula Injection (OWASP): se o texto começar com um
    caractere que Excel/Google Sheets interpretam como início de fórmula,
    prefixa com apóstrofo para forçar leitura como texto literal — mesmo
    efeito de digitar ' antes do valor manualmente na célula."""
    s = "" if valor is None else str(valor)
    if s and s[0] in _CARACTERES_INICIO_FORMULA:
        return "'" + s
    return s

def _pasta_esta_em_sincronizacao_nuvem():
    """Detecta, de forma best-effort, se a pasta do projeto está dentro de
    uma pasta sincronizada pelo OneDrive nesta máquina — isso poderia
    enviar dados pessoais/sigilosos (LGPD, PLD/FT) e credenciais para a
    nuvem corporativa sem controle adicional, mesmo estando fora do Git."""
    caminho_projeto = os.path.abspath(os.path.dirname(__file__)).lower()
    for var in ("OneDrive", "OneDriveCommercial", "OneDriveConsumer"):
        base = os.environ.get(var)
        if base and caminho_projeto.startswith(os.path.abspath(base).lower()):
            return True
    return False

# ==========================================================
# EXPORTAR / IMPORTAR LOTE ENTRE ANALISTAS (mesclar duas máquinas)
# ==========================================================
# lote_casos.json fica criptografado por DPAPI e amarrado à conta Windows
# de quem o gravou — por isso não pode ser copiado direto de uma máquina
# para outra. Este export/import gera um arquivo PORTÁTIL separado (sem
# senha — trocado só entre as duas pessoas do time por canal já confiável
# da empresa, ex. Teams/e-mail corporativo).

def exportar_lote_arquivo(casos):
    return json.dumps(casos, ensure_ascii=False).encode("utf-8")

def importar_lote_arquivo(bytes_arquivo):
    return json.loads(bytes_arquivo.decode("utf-8"))

def mesclar_lotes_bloqueando_colisao(casos_locais, casos_externos):
    """Mescla o lote local com um lote importado de outro analista. Se
    algum numero_caso aparecer nos dois lados, a mesclagem é BLOQUEADA
    (não decide sozinho qual dos dois manter) — retorna (None, colisões)
    para o analista resolver manualmente antes de tentar de novo."""
    por_numero_local = {c["numero_caso"]: c for c in casos_locais}
    por_numero_externo = {c["numero_caso"]: c for c in casos_externos}
    colisoes = sorted(set(por_numero_local) & set(por_numero_externo))
    if colisoes:
        detalhes = [
            {
                "numero_caso": n,
                "nome_local": por_numero_local[n].get("nome", ""),
                "nome_no_arquivo_importado": por_numero_externo[n].get("nome", ""),
            }
            for n in colisoes
        ]
        return None, detalhes
    mesclado = sorted(casos_locais + casos_externos, key=lambda c: c["numero_caso"])
    return mesclado, []

if not DPAPI_DISPONIVEL:
    st.warning(
        "O módulo **pywin32** não está instalado nesta máquina — historico_casos.json, "
        "lote_casos.json e token_google.json estão sendo salvos SEM a criptografia adicional "
        "(DPAPI). Rode `py -m pip install -r requirements.txt` para habilitar.",
        icon=":material/warning:",
    )

if _pasta_esta_em_sincronizacao_nuvem():
    st.warning(
        "Esta pasta do projeto parece estar dentro de um diretório sincronizado pelo "
        "OneDrive nesta máquina. Os arquivos locais contêm dados pessoais e sigilosos de "
        "PLD/FT e credenciais — confirme com a TI se essa sincronização deve ficar ativa "
        "para esta pasta específica.",
        icon=":material/warning:",
    )

# ==========================================================
# CONFIGURAÇÃO DO LOTE XML SISCOAF
# ==========================================================

# CNPJ da comunicante (somente números). Definir em COAF_CNPJ_COMUNICANTE; o padrão é um placeholder.
CNPJ_COMUNICANTE = os.environ.get("COAF_CNPJ_COMUNICANTE", "00000000000000")

# Segmento/modelo confirmados na tela "Registrar Comunicação" do SISCOAF:
# "SPA - Apostas de quota fixa" (único, não há outras opções para o setor)

# Códigos de enquadramento (CodEnq) confirmados na tela real do SISCOAF
# (tabela "Selecione as ocorrências", códigos 1407 a 1427, todos ligados
# a artigos da Portaria SPA/MF 1.143/2024). Mapeamento por gatilho interno:
CODENQ_MAP = {
    "incompatibilidade_financeira": 1418,  # Art. 25-IX (incompatibilidade renda/operações)
    "fracionamento": 1420,                  # Art. 25-XI (aporte/retirada em curto tempo)
    "pep": 1426,                            # Art. 25-XVII (conta aberta em nome de PEP)
    "aporte_origem_suspeita": 1415,         # Art. 25-VI (origem suspeita dos valores)
    # Antecedente criminal GENÉRICO (sem relação identificada com crime
    # financeiro) — sem código específico confirmado, fica só como texto
    # no campo Det.
    "antecedente_criminal": None,
    # Antecedente criminal ESPECIFICAMENTE ligado a lavagem de dinheiro ou
    # crime contra o sistema financeiro (ex: corrupção, fraude, lavagem em
    # processo à parte) — confirmado na tabela real do SISCOAF.
    "antecedente_crime_financeiro": 1410,  # Art. 25-I
    "reporte_anterior": None,        # TODO: não identificado código dedicado
}

# TpEnv fixo em "Titular" (1) — única opção visível no formulário do
# segmento SPA. Se um dia for necessário reportar um terceiro concreto
# (ex: destinatário de saque via Pix), revisar este ponto.
TPENV_TITULAR = 1

# ==========================================================
# FAIXAS DE NÚMERO DE CASO POR ANALISTA
# ==========================================================
# A numeração de caso NÃO é uma sequência contínua para sempre — ela
# reinicia a cada lote enviado ao SISCOAF (na prática, ~mensalmente,
# junto com o documento mestre do mês). Dentro de um mesmo lote/mês,
# porém, cada analista numera os PRÓPRIOS casos dentro de uma faixa fixa
# e exclusiva, para que duas pessoas trabalhando em máquinas separadas
# NUNCA gerem o mesmo número de caso naquele ciclo — o que causaria
# conflito ao mesclar os lotes das duas máquinas em um só (ver
# "🗑️ Limpar lote" na sidebar, que reinicia o contador de todo mundo
# para o início da própria faixa quando o envio é confirmado).
#
# Esse número é só controle interno (aparece no texto do reporte e no
# histórico local); o SISCOAF não exige nenhum valor específico aqui —
# o NumOcorrencia do XML é renumerado do zero a cada exportação,
# independente disso (ver gerar_xml_lote_bytes).
FAIXA_NUMERO_POR_ANALISTA = {
    "Analista A": 10000,
    "Analista B": 20000,
    "Outro": 90000,
}
TAMANHO_FAIXA_POR_ANALISTA = 10000  # cada analista tem até 9999 casos na própria faixa

def proximo_numero_para_analista(historico, analista_nome):
    """Calcula o próximo número de caso dentro da faixa exclusiva do
    analista informado, a partir do que já foi salvo no histórico local."""
    inicio_faixa = FAIXA_NUMERO_POR_ANALISTA.get(analista_nome, FAIXA_NUMERO_POR_ANALISTA["Outro"])
    contadores = historico.get("proximo_numero_por_analista", {})
    return contadores.get(analista_nome, inicio_faixa + 1)

def numero_esta_na_faixa_do_analista(numero, analista_nome):
    inicio_faixa = FAIXA_NUMERO_POR_ANALISTA.get(analista_nome, FAIXA_NUMERO_POR_ANALISTA["Outro"])
    return inicio_faixa < numero <= inicio_faixa + TAMANHO_FAIXA_POR_ANALISTA

# ⚠️ HIPÓTESE AINDA NÃO CONFIRMADA COM O SISCOAF — mapeamento dos 3 campos
# de valor visíveis na tela ("Valor do(s) prêmio(s)", "Valor da(s) aposta(s)",
# "Valor de operação de outro tipo") para as tags reais do XML (VlCred/VlDeb/
# VlProv/VlProp). Testar no ambiente de treinamento antes de usar em produção.
#   VlCred = Valor do(s) prêmio(s)      -> hipótese: total sacado (prêmios pagos)
#   VlDeb  = Valor da(s) aposta(s)       -> hipótese: total depositado (apostado)
#   VlProv = (não identificado na tela) -> fixado em 0
#   VlProp = Valor de operação de outro tipo -> fixado em 0 por padrão


# Autenticação via OAuth da SUA conta Google (não conta de serviço).
# Contorna a política da organização que bloqueia compartilhar planilhas
# com contas externas: aqui o acesso é o mesmo que você já tem na planilha,
# não precisa compartilhar nada com ninguém.
OAUTH_CLIENT_FILE = os.path.join(os.path.dirname(__file__), "oauth_client.json")
TOKEN_FILE = os.path.join(os.path.dirname(__file__), "token_google.json")

# Escopo mínimo necessário: só Sheets (leitura/escrita). Não pedimos acesso
# geral ao Drive nem a outros serviços do Google.
GOOGLE_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

@st.cache_resource
def conectar_sheets():
    """Autentica com a conta Google da pessoa (fluxo OAuth). Na primeira vez,
    abre o navegador pedindo login/autorização; nas próximas, reutiliza o
    token salvo localmente (token_google.json) sem pedir login de novo."""
    if not GSPREAD_DISPONIVEL:
        return None, "As bibliotecas gspread/google-auth-oauthlib não estão instaladas."
    if not os.path.exists(OAUTH_CLIENT_FILE):
        return None, f"Arquivo de credenciais OAuth não encontrado em: {OAUTH_CLIENT_FILE}"
    _restringir_acesso_arquivo(OAUTH_CLIENT_FILE)

    creds = None
    if os.path.exists(TOKEN_FILE):
        try:
            token_info = _ler_json_protegido(TOKEN_FILE, None, "coaf_token_google")
            if token_info:
                creds = UserCredentials.from_authorized_user_info(token_info, GOOGLE_SCOPES)
        except Exception:
            creds = None

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(GoogleAuthRequest())
            except Exception:
                creds = None
        if not creds:
            try:
                flow = InstalledAppFlow.from_client_secrets_file(OAUTH_CLIENT_FILE, GOOGLE_SCOPES)
                creds = flow.run_local_server(port=0)
            except Exception as e:
                return None, f"Falha na autenticação Google: {e}"
        _salvar_json_protegido(TOKEN_FILE, json.loads(creds.to_json()), "coaf_token_google")

    try:
        client = gspread.authorize(creds)
        return client, None
    except Exception as e:
        return None, f"Falha ao autorizar o cliente do Sheets: {e}"

def enviar_para_sheets(sheet_id, aba, linha_valores):
    """Adiciona uma linha na planilha. Retorna (sucesso: bool, mensagem: str)."""
    client, erro = conectar_sheets()
    if erro:
        return False, erro
    try:
        planilha = client.open_by_key(sheet_id)
        worksheet = planilha.worksheet(aba)
        worksheet.append_row(linha_valores, value_input_option="USER_ENTERED")
        return True, "Linha adicionada com sucesso."
    except gspread.exceptions.APIError as e:
        return False, f"Erro de API do Google (confira se a planilha foi compartilhada com a conta de serviço): {e}"
    except Exception as e:
        return False, f"Erro ao enviar para o Sheets: {e}"

# ========================
# FUNÇÕES AUXILIARES
# ========================

def carregar_historico():
    return _ler_json_protegido(
        HISTORICO_FILE,
        {"casos": [], "proximo_numero": 1, "proximo_numero_por_analista": {}},
        "coaf_historico_casos",
    )

def salvar_historico(historico):
    _salvar_json_protegido(HISTORICO_FILE, historico, "coaf_historico_casos")

def parse_valor(valor_str):
    if not valor_str:
        return 0.0
    s = str(valor_str).replace("R$", "").replace(" ", "").strip()
    # Detectar sinal negativo
    negativo = s.startswith("-") or s.startswith("(")
    s = s.replace("-", "").replace("(", "").replace(")", "")
    s = s.replace(".", "").replace(",", ".")
    try:
        v = float(s)
        return -v if negativo else v
    except:
        return 0.0

def formatar_valor(valor):
    return f"R$ {abs(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def calcular_multiplo(total, renda):
    if renda == 0:
        return 0
    return round(abs(total) / renda)

def calcular_ticket_medio(total, qtd):
    if qtd == 0:
        return 0
    return abs(total) / qtd

def calcular_percentual(valor, total):
    if total == 0:
        return 0
    return round((abs(valor) / abs(total)) * 100, 1)

def formatar_pct(valor):
    """Formata percentual com vírgula decimal (padrão BR), evitando o bug de
    str(41.5) sair com ponto em vez de vírgula ("41,5")."""
    return str(valor).replace(".", ",")

def formatar_cpf(cpf):
    cpf = re.sub(r'\D', '', str(cpf))
    if len(cpf) == 11:
        return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"
    return cpf

def formatar_milhar(qtd):
    """Formata um inteiro com separador de milhar no padrão BR (ponto),
    isolado de qualquer outra formatação de valor monetário — usado para
    não contaminar acidentalmente números com vírgula decimal já formatados
    (bug corrigido: antes um .replace(",", ".") genérico quebrava o R$ de
    valores monetários que já estavam corretos)."""
    return f"{qtd:,}".replace(",", ".")

# Banco de aberturas alternativas para o parágrafo de fundamentação, usado
# para variar a redação entre casos (sorteio aleatório a cada reporte).
ABERTURAS_FUNDAMENTACAO = [
    "Os dados sob análise confirmam",
    "Os dados apurados evidenciam",
    "Os dados levantados demonstram",
    "Os dados obtidos apontam para",
    "Os dados consolidados revelam",
    "As informações sob análise confirmam",
    "As informações apuradas evidenciam",
    "As informações levantadas demonstram",
    "As informações coletadas apontam para",
    "As informações obtidas revelam",
    "Os elementos apurados confirmam",
    "Os elementos identificados evidenciam",
    "Os elementos analisados demonstram",
    "Os elementos levantados apontam para",
    "Os registros analisados confirmam",
    "Os registros obtidos evidenciam",
    "Os registros levantados demonstram",
    "Os achados da análise confirmam",
    "Os achados apurados evidenciam",
    "Os achados identificados demonstram",
    "O conjunto dos dados analisados confirma",
    "A análise dos registros evidencia",
    "O exame das informações levantadas demonstra",
    "A verificação dos dados obtidos aponta para",
    "O cruzamento das informações apuradas revela",
]

FRASES_COMPLETAS_FUNDAMENTACAO = [
    "As informações apuradas evidenciam movimentações financeiras atípicas na conta {art}, cujos montantes aportados na plataforma não apresentam compatibilidade com sua capacidade financeira presumida.",
]

CONTINUACOES_ABERTURA_FUNDAMENTACAO = [
    "movimentações financeiras atípicas na conta {art}, cujos valores transacionados não apresentam compatibilidade com sua capacidade financeira presumida.",
    "um padrão de movimentação financeira incompatível com a capacidade econômica presumida {art}.",
    "operações financeiras cujo volume não guarda relação com os rendimentos declarados ou presumidos {art}.",
    "um fluxo de recursos incompatível com o perfil financeiro presumido {art}.",
]

FECHAMENTOS_FUNDAMENTACAO = [
    "Nesse contexto, não se pode afastar a hipótese de que os recursos movimentados sejam de titularidade de terceiros ou de origem não justificável por fontes lícitas identificadas.",
    "Diante do exposto, não se pode descartar a possibilidade de que os valores movimentados pertençam a terceiros ou tenham origem não identificada em fontes lícitas.",
    "Ante o contexto apresentado, subsiste a hipótese de que os recursos em questão sejam de titularidade alheia ou provenientes de origem não comprovadamente lícita.",
    "Assim, não se afasta a possibilidade de que os montantes movimentados tenham origem em recursos de terceiros ou em fontes cuja licitude não foi identificada.",
    "Nesses termos, permanece a hipótese de titularidade de terceiros sobre os recursos movimentados, ou de origem não lícita não identificada.",
]

def periodo_texto_para_datas(periodo_txt):
    """Converte um texto tipo 'MÊS 06/2026' no primeiro e último dia daquele
    mês (como objetos date), para pré-preencher os campos de data do lote
    XML sem precisar digitar a mesma informação duas vezes. Se não conseguir
    interpretar o texto, usa a data de hoje como padrão."""
    m = re.search(r'(\d{1,2})/(\d{4})', periodo_txt or "")
    if not m:
        hoje = date.today()
        return hoje, hoje
    mes, ano = int(m.group(1)), int(m.group(2))
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    return date(ano, mes, 1), date(ano, mes, ultimo_dia)

MESES_NOME = {
    "01": "janeiro", "02": "fevereiro", "03": "março", "04": "abril",
    "05": "maio", "06": "junho", "07": "julho", "08": "agosto",
    "09": "setembro", "10": "outubro", "11": "novembro", "12": "dezembro",
}

def formatar_periodo_extenso(mes, ano):
    """Converte mês (01-12) e ano em 'mês/ano' por extenso, ex: 'fevereiro/2026'."""
    return f"{MESES_NOME.get(str(mes).zfill(2), mes)}/{ano}"

MAX_EMPRESAS = 5
MAX_PERIODOS_ANTERIORES = 5
CAMPOS_PARA_LIMPAR = {
    "cpf_busca_input": "",
    "data_nasc_raw": "",
    "cidade_uf_input": "",
    "pep_input": "NÃO",
    "genero_opcao_input": "Automático",
    "renda_raw": "",
    "tem_cnpj_input": False,
    "qtd_empresas_input": 1,
    "antecedente_input": False,
    "antecedente_crime_fin_input": False,
    "detalhe_crim_input": "",
    "reporte_anterior_input": False,
    "qtd_periodos_ant_input": 1,
    "bpc_input": False,
    "detalhe_bpc_input": "",
    "pep_relacionado_input": False,
    "detalhe_pep_input": "",
    "negativado_input": False,
    "valor_neg_input": "",
    "obs_input": "",
    "lavagem_input": True,
    "financiamento_input": False,
    "manipulacao_input": False,
}
for _i in range(MAX_EMPRESAS):
    CAMPOS_PARA_LIMPAR.update({
        f"cnpj_emp_{_i}": "",
        f"tipo_empresa_{_i}": "MEI",
        f"fat_{_i}": "",
        f"participacao_{_i}": "",
        f"ramo_{_i}": "",
        f"inapta_{_i}": False,
    })
for _i in range(MAX_PERIODOS_ANTERIORES):
    CAMPOS_PARA_LIMPAR.update({
        f"mes_ant_{_i}": "01",
        f"ano_ant_{_i}": 2026,
        f"dep_ant_{_i}": "",
        f"saq_ant_{_i}": "",
    })

def limpar_campos_do_formulario():
    for chave, valor_padrao in CAMPOS_PARA_LIMPAR.items():
        st.session_state[chave] = valor_padrao

LABELS_SUBLINHAR = ["Valor(es) envolvido(s):", "Análise interna:", "Fundamentação legal:", "Classificação da suspeita:"]

def gerar_docx_bytes(texto, fonte="Montserrat", tamanho_pt=11):
    doc = Document()

    estilo_normal = doc.styles['Normal']
    estilo_normal.font.name = fonte
    estilo_normal.font.size = Pt(tamanho_pt)
    rpr = estilo_normal.element.get_or_add_rPr()
    from docx.oxml.ns import qn
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = rpr.makeelement(qn('w:rFonts'), {})
        rpr.append(rfonts)
    rfonts.set(qn('w:eastAsia'), fonte)

    for linha in texto.split('\n'):
        label_encontrado = next((lb for lb in LABELS_SUBLINHAR if linha.startswith(lb)), None)

        p = doc.add_paragraph()
        if label_encontrado:
            run_label = p.add_run(label_encontrado)
            run_label.font.name = fonte
            run_label.font.size = Pt(tamanho_pt)
            run_label.underline = True
            resto = linha[len(label_encontrado):]
            if resto:
                run_resto = p.add_run(resto)
                run_resto.font.name = fonte
                run_resto.font.size = Pt(tamanho_pt)
        else:
            run = p.add_run(linha)
            run.font.name = fonte
            run.font.size = Pt(tamanho_pt)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def _mask_data_nasc():
    raw = re.sub(r'\D', '', st.session_state.get('data_nasc_raw', ''))[:8]
    if len(raw) >= 5:
        formatted = f"{raw[:2]}/{raw[2:4]}/{raw[4:]}"
    elif len(raw) >= 3:
        formatted = f"{raw[:2]}/{raw[2:]}"
    else:
        formatted = raw
    st.session_state['data_nasc_raw'] = formatted

def _mask_renda():
    val = st.session_state.get('renda_raw', '')

    if ',' in val:
        inteiro_part, _, dec_part = val.partition(',')
        inteiro_digits = re.sub(r'\D', '', inteiro_part)
        dec_digits = re.sub(r'\D', '', dec_part)[:2].ljust(2, '0')
        inteiro_digits = str(int(inteiro_digits)) if inteiro_digits else '0'
    else:
        inteiro_digits = re.sub(r'\D', '', val)
        if not inteiro_digits:
            st.session_state['renda_raw'] = ''
            return
        inteiro_digits = str(int(inteiro_digits))
        dec_digits = '00'

    partes = []
    while len(inteiro_digits) > 3:
        partes.insert(0, inteiro_digits[-3:])
        inteiro_digits = inteiro_digits[:-3]
    partes.insert(0, inteiro_digits)
    st.session_state['renda_raw'] = f"{'.'.join(partes)},{dec_digits}"

def extrair_linha_pdf(pdf_file, cpf_busca):
    cpf_limpo = re.sub(r'\D', '', cpf_busca)
    try:
        pdf_file.seek(0)
    except Exception:
        pass

    try:
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables({
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "lines",
                    "text_x_tolerance": 1.5,
                })

                if not tables:
                    tables = page.extract_tables({"text_x_tolerance": 1.5})

                for table in tables:
                    if not table or len(table) < 2:
                        continue

                    header_idx = None
                    for i, row in enumerate(table):
                        if row and any(str(c).upper().strip() in ['CPF', 'NOME', 'NOME COMPLETO'] for c in row if c):
                            header_idx = i
                            break

                    if header_idx is None:
                        continue

                    raw_headers = table[header_idx]
                    headers = [str(h).strip().upper() if h else '' for h in raw_headers]

                    def find_col(keywords, start=0, end=None):
                        end = end or len(headers)
                        for i in range(start, end):
                            for kw in keywords:
                                if kw in headers[i]:
                                    return i
                        return None

                    NUM_HEADERS = {'Nº', 'N°', 'N.', 'NUM', 'NUMERO', 'NÚMERO'}
                    def find_num_col(value_idx, span=5):
                        if value_idx is None:
                            return None
                        end = min(value_idx + 1 + span, len(headers))
                        for i in range(value_idx + 1, end):
                            if headers[i].strip() in NUM_HEADERS:
                                return i
                        return None

                    cpf_idx = find_col(['CPF'])
                    nome_idx = find_col(['NOME'])
                    dep_idx = find_col(['DEPÓSITO', 'DEPOSITO', 'DEP.'])
                    n_dep_idx = find_num_col(dep_idx)
                    saq_idx = find_col(['SAQUE', 'SAQ.'])
                    n_saq_idx = find_num_col(saq_idx)
                    cas_idx = find_col(['CASSINO', 'CASINO', 'CASS'])
                    n_cas_idx = find_num_col(cas_idx, span=3)
                    sb_idx = find_col(['SB', 'SPORT', 'ESPORT'])
                    n_sb_idx = find_num_col(sb_idx, span=3)

                    for row in table[header_idx + 1:]:
                        if not row:
                            continue

                        cpf_cell = str(row[cpf_idx]).strip() if cpf_idx is not None and cpf_idx < len(row) else ""
                        cpf_linha = re.sub(r'\D', '', cpf_cell)

                        if cpf_linha != cpf_limpo:
                            continue

                        row_clean = [str(c).strip() if c else '' for c in row]

                        nome_raw = row_clean[nome_idx] if nome_idx and nome_idx < len(row_clean) else ''
                        if nome_idx and nome_idx + 1 < len(row_clean):
                            prox = row_clean[nome_idx + 1]
                            if prox and not re.match(r'^R\$|^\d', prox) and '@' not in prox:
                                nome_raw = nome_raw + ' ' + prox
                        nome_raw = ' '.join(nome_raw.split())

                        def get_val(idx):
                            if idx is not None and idx < len(row_clean):
                                return parse_valor(row_clean[idx])
                            return 0.0

                        def get_int(idx):
                            v = get_val(idx)
                            return int(abs(v))

                        return {
                            'nome': nome_raw,
                            'cpf': formatar_cpf(cpf_limpo),
                            'total_dep': get_val(dep_idx),
                            'qtd_dep': get_int(n_dep_idx),
                            'total_saq': get_val(saq_idx),
                            'qtd_saq': get_int(n_saq_idx),
                            'val_cas': get_val(cas_idx),
                            'qtd_cas': get_int(n_cas_idx),
                            'val_sb': get_val(sb_idx),
                            'qtd_sb': get_int(n_sb_idx),
                        }
    except Exception as e:
        st.warning(f"Erro ao processar PDF: {e}")

    return None

@st.cache_data(show_spinner=False)
def buscar_cpf_em_pdfs(pdf_files, cpf_busca):
    resultado = []
    for pdf_file in pdf_files:
        dados = extrair_linha_pdf(pdf_file, cpf_busca)
        if dados:
            resultado.append({
                'casa': pdf_file.name.replace('.pdf', '').replace('_', ' '),
                **dados
            })
    return resultado


# ==========================================================
# FUNÇÕES DO LOTE XML (SISCOAF)
# ==========================================================

def carregar_lote_casos():
    return _ler_json_protegido(LOTE_CASOS_FILE, [], "coaf_lote_casos")

def salvar_lote_casos(casos):
    _salvar_json_protegido(LOTE_CASOS_FILE, casos, "coaf_lote_casos")

def carregar_arquivo_casos():
    return _ler_json_protegido(ARQUIVO_CASOS_FILE, [], "coaf_arquivo_casos")

def salvar_arquivo_casos(casos):
    _salvar_json_protegido(ARQUIVO_CASOS_FILE, casos, "coaf_arquivo_casos")

def _chave_caso(caso):
    return (caso.get("dt_inicio"), caso.get("dt_fim"), caso.get("numero_caso"))

def arquivar_caso(caso):
    """Guarda uma cópia PERMANENTE do caso (texto completo incluído) no
    arquivo histórico — nunca é apagada por 'Limpar lote'. Identifica o
    caso pela combinação (dt_inicio, dt_fim, numero_caso): reprocessar o
    mesmo caso dentro do mesmo ciclo atualiza a cópia arquivada em vez de
    duplicar (numero_caso sozinho não é uma chave segura entre ciclos,
    já que reinicia a cada mês).

    Todo caso entra como status "rascunho" (editável/excluível na tela de
    consulta). Um caso já marcado "enviado" (ver marcar_casos_como_enviados)
    NUNCA é sobrescrito aqui — trava contra alterar retroativamente o
    registro do que já foi de fato comunicado ao COAF."""
    arquivo = carregar_arquivo_casos()
    chave = _chave_caso(caso)
    if any(_chave_caso(c) == chave and c.get("status") == "enviado" for c in arquivo):
        return
    arquivo = [c for c in arquivo if _chave_caso(c) != chave]
    caso_arquivado = dict(caso)
    caso_arquivado["status"] = "rascunho"
    caso_arquivado["arquivado_em"] = datetime.now().strftime("%d/%m/%Y %H:%M")
    arquivo.append(caso_arquivado)
    salvar_arquivo_casos(arquivo)

def marcar_casos_como_enviados(casos):
    """Trava no arquivo permanente os casos que acabaram de ser incluídos
    num lote confirmado (ver botão 'Limpar lote'): a partir daqui, status
    vira "enviado" e a tela de consulta para de oferecer editar/excluir
    para eles — preserva o registro fiel do que foi comunicado ao COAF."""
    arquivo = carregar_arquivo_casos()
    chaves_enviadas = {_chave_caso(c) for c in casos}
    agora = datetime.now().strftime("%d/%m/%Y %H:%M")
    for item in arquivo:
        if _chave_caso(item) in chaves_enviadas:
            item["status"] = "enviado"
            item["enviado_em"] = agora
    salvar_arquivo_casos(arquivo)

def editar_caso_arquivado(dt_inicio, dt_fim, numero_caso, novo_texto):
    """Atualiza o texto de um caso ainda em rascunho. Levanta ValueError
    se o caso já estiver marcado como enviado (travado)."""
    chave = (dt_inicio, dt_fim, numero_caso)
    arquivo = carregar_arquivo_casos()
    encontrado = False
    for c in arquivo:
        if _chave_caso(c) == chave:
            if c.get("status") == "enviado":
                raise ValueError("Este caso já foi enviado e está travado para edição.")
            c["det"] = novo_texto
            c["editado_em"] = datetime.now().strftime("%d/%m/%Y %H:%M")
            encontrado = True
    if encontrado:
        salvar_arquivo_casos(arquivo)

    # se ainda estiver na fila do lote atual (mesmo ciclo, ainda não
    # enviado), mantém os dois em sincronia
    lote = carregar_lote_casos()
    mudou_lote = False
    for c in lote:
        if _chave_caso(c) == chave:
            c["det"] = novo_texto
            mudou_lote = True
    if mudou_lote:
        salvar_lote_casos(lote)

def excluir_caso_arquivado(dt_inicio, dt_fim, numero_caso):
    """Remove um caso em rascunho do arquivo permanente e, se ainda
    estiver na fila do lote atual, remove de lá também. Levanta
    ValueError se o caso já estiver marcado como enviado (travado)."""
    chave = (dt_inicio, dt_fim, numero_caso)
    arquivo = carregar_arquivo_casos()
    alvo = next((c for c in arquivo if _chave_caso(c) == chave), None)
    if alvo and alvo.get("status") == "enviado":
        raise ValueError("Este caso já foi enviado e está travado para exclusão.")
    salvar_arquivo_casos([c for c in arquivo if _chave_caso(c) != chave])
    salvar_lote_casos([c for c in carregar_lote_casos() if _chave_caso(c) != chave])

def adicionar_caso_ao_lote(dados_caso):
    casos = carregar_lote_casos()
    casos.append(dados_caso)
    salvar_lote_casos(casos)
    arquivar_caso(dados_caso)

def limpar_lote_casos():
    salvar_lote_casos([])

def periodo_para_datas(periodo_txt):
    """Converte 'MÊS 06/2026' em (DtInicio, DtFim) dd/mm/aaaa, usando o
    primeiro e o último dia do mês informado."""
    m = re.search(r'(\d{1,2})/(\d{4})', periodo_txt)
    if not m:
        return "", ""
    mes, ano = int(m.group(1)), int(m.group(2))
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    return f"01/{mes:02d}/{ano}", f"{ultimo_dia:02d}/{mes:02d}/{ano}"

def montar_enquadramentos(pep_flag, fracionamento_flag, aporte_suspeito_flag=True,
                           antecedente_crime_financeiro_flag=False):
    """Monta a lista de CodEnq aplicáveis a um caso, a partir dos gatilhos
    já identificados. Ignora entradas None do CODENQ_MAP (pendências)."""
    codigos = []
    if aporte_suspeito_flag:
        codigos.append(CODENQ_MAP["aporte_origem_suspeita"])
    codigos.append(CODENQ_MAP["incompatibilidade_financeira"])
    if pep_flag:
        codigos.append(CODENQ_MAP["pep"])
    if fracionamento_flag:
        codigos.append(CODENQ_MAP["fracionamento"])
    if antecedente_crime_financeiro_flag:
        codigos.append(CODENQ_MAP["antecedente_crime_financeiro"])
    return [c for c in codigos if c is not None]

def sanitizar_texto_siscoaf(texto):
    """Remove/substitui caracteres que o SISCOAF rejeita no campo Det e dados cadastrais
    (erro 525: 'não pode conter os caracteres <, >, &, ", ', |').
    Também normaliza aspas curvas, apóstrofos tipográficos, travessões e símbolos fora do ISO-8859-1."""
    if not texto:
        return ""
    texto = str(texto)
    # Substituições de caracteres expressamente proibidos pelo SISCOAF
    texto = texto.replace('&', ' e ')
    texto = texto.replace('"', '')
    texto = texto.replace("'", ' ')  # espaço no lugar do apóstrofo (D'Oeste -> D Oeste)
    texto = texto.replace('<', '').replace('>', '').replace('|', '')
    # Aspas e apóstrofos tipográficos
    texto = texto.replace('“', '').replace('”', '').replace('‘', ' ').replace('’', ' ')
    # Travessões e hífens tipográficos
    texto = texto.replace('—', '-').replace('–', '-')
    # Símbolos gráficos especiais
    texto = texto.replace('⛝', '[X]').replace('☐', '[ ]')
    # Normaliza espaços múltiplos
    texto = re.sub(r'[ \t]{2,}', ' ', texto)
    return texto.strip()

def garantir_iso_8859_1(texto):
    """Garante que a string só contenha caracteres suportados por ISO-8859-1,
    evitando que o ElementTree gere entidades numéricas XML contendo '&' (ex: &#9744;),
    o que aciona o erro 525 de caractere proibido no validador do SISCOAF."""
    if not texto:
        return ""
    try:
        texto.encode("iso-8859-1")
        return texto
    except UnicodeEncodeError:
        import unicodedata
        normalizado = unicodedata.normalize('NFKD', texto)
        return normalizado.encode("iso-8859-1", "ignore").decode("iso-8859-1")

def gerar_xml_lote_bytes():
    """Monta o LOTE inteiro a partir de todos os casos acumulados em
    lote_casos.json, no leiaute oficial do SISCOAF. Retorna bytes prontos
    para download.

    ATENÇÃO: o mapeamento de VlCred/VlDeb/VlProv/VlProp ainda é uma
    HIPÓTESE não confirmada — testar no ambiente de treinamento do
    SISCOAF antes de qualquer envio em produção.

    NumOcorrencia É RENUMERADO DO ZERO (1, 2, 3, ...) a cada exportação,
    na ordem dos casos do lote — é só um identificador exigido pelo
    SISCOAF dentro deste arquivo específico (não precisa, e não deve,
    ser o mesmo "número de caso" usado no controle interno/relatório,
    que segue sua própria faixa por analista em `numero_caso`)."""
    casos = carregar_lote_casos()
    if not casos:
        return None

    casos_ordenados = sorted(casos, key=lambda c: c["numero_caso"])

    lote_id = f"SISCOAF{datetime.now().strftime('%d%m%Y')}"

    root = ET.Element("LOTE")
    ocorrencias_el = ET.SubElement(root, "OCORRENCIAS", ID=lote_id)

    for num_ocorrencia, caso in enumerate(casos_ordenados, start=1):
        ocorrencia_el = ET.SubElement(ocorrencias_el, "OCORRENCIA")
        ET.SubElement(ocorrencia_el, "CPFCNPJCom").text = CNPJ_COMUNICANTE
        ET.SubElement(ocorrencia_el, "NumOcorrencia").text = str(num_ocorrencia)
        ET.SubElement(ocorrencia_el, "DtInicio").text = caso["dt_inicio"]
        ET.SubElement(ocorrencia_el, "DtFim").text = caso["dt_fim"]
        ET.SubElement(ocorrencia_el, "AgMun").text = garantir_iso_8859_1(sanitizar_texto_siscoaf(caso["ag_mun"]))
        ET.SubElement(ocorrencia_el, "AgUF").text = caso["ag_uf"].strip().upper()[:2]
        # HIPÓTESE: VlCred = prêmios (aqui aproximado por total sacado),
        # VlDeb = apostas (aqui aproximado por total depositado). Confirmar.
        ET.SubElement(ocorrencia_el, "VlCred").text = str(caso["vl_cred"])
        ET.SubElement(ocorrencia_el, "VlDeb").text = str(caso["vl_deb"])
        ET.SubElement(ocorrencia_el, "VlProv").text = str(caso.get("vl_prov", 0))
        ET.SubElement(ocorrencia_el, "VlProp").text = str(caso.get("vl_prop", 0))
        ET.SubElement(ocorrencia_el, "Det").text = garantir_iso_8859_1(sanitizar_texto_siscoaf(caso["det"]))

        enq_el = ET.SubElement(ocorrencia_el, "ENQUADRAMENTOS")
        for codigo in caso["enquadramentos"]:
            ET.SubElement(enq_el, "CodEnq").text = str(codigo)

        env_wrapper = ET.SubElement(ocorrencia_el, "ENVOLVIDOS")
        env_el = ET.SubElement(env_wrapper, "ENVOLVIDO")
        ET.SubElement(env_el, "CPFCNPJEnv").text = caso["cpf"]
        ET.SubElement(env_el, "NmEnv").text = garantir_iso_8859_1(sanitizar_texto_siscoaf(caso["nome"]))
        ET.SubElement(env_el, "TpEnv").text = str(TPENV_TITULAR)
        ET.SubElement(env_el, "PObrigada").text = "0"
        ET.SubElement(env_el, "PEP").text = str(caso.get("pep_flag", 0))
        ET.SubElement(env_el, "ServPub").text = "0"

    xml_bytes = ET.tostring(root, encoding="iso-8859-1", xml_declaration=True)
    return xml_bytes


# ========================
# CARREGAR HISTÓRICO
# ========================
historico = carregar_historico()

# ========================
# SIDEBAR - HISTÓRICO E LOTE
# ========================
with st.sidebar:
    st.header("📊 Controle de casos")
    st.caption("Próximo número por analista (faixa fixa evita colisão ao mesclar; "
               "reinicia a cada lote enviado, não é contínuo entre meses):")
    for _analista_nome in FAIXA_NUMERO_POR_ANALISTA:
        st.metric(f"Próximo nº — {_analista_nome}", proximo_numero_para_analista(historico, _analista_nome))
    st.metric("Total de casos registrados", len(historico['casos']))

    if historico['casos']:
        st.markdown("**Últimos casos:**")
        for caso in historico['casos'][-5:][::-1]:
            st.markdown(f"- **{caso['numero']}** — {caso['nome'][:25]}...")

    if st.button("🗑️ Limpar histórico", type="secondary"):
        historico = {"casos": [], "proximo_numero": 1, "proximo_numero_por_analista": {}}
        salvar_historico(historico)
        st.rerun()

    st.markdown("---")
    st.header("📦 Lote SISCOAF")
    lote_atual = carregar_lote_casos()
    st.metric("Casos no lote atual", len(lote_atual))

    if lote_atual:
        xml_lote = gerar_xml_lote_bytes()
        st.download_button(
            label="⬇️ Baixar Lote XML",
            data=xml_lote,
            file_name=f"lote_coaf_{datetime.now().strftime('%Y%m%d')}.xml",
            mime="application/xml",
            use_container_width=True,
        )
        st.caption(":material/warning: Teste este arquivo em 'Validar Arquivo' no ambiente de "
                   "treinamento do SISCOAF (treina.siscoaf.coaf.gov.br) antes "
                   "de usar em produção — o mapeamento de valores ainda é hipótese.")
        st.caption("Ao confirmar o envio, a numeração de caso de cada analista volta para o "
                   "início da própria faixa (não é contínua entre lotes/meses). Os casos deste "
                   "lote também ficam travados no arquivo permanente (não dá mais pra editar/excluir).")
        confirmar_limpeza_lote = st.checkbox(
            "Confirmo que o lote já foi enviado ao SISCOAF", key="confirma_limpar_lote"
        )
        if st.button("🗑️ Limpar lote (após envio confirmado)", type="secondary",
                     disabled=not confirmar_limpeza_lote):
            marcar_casos_como_enviados(lote_atual)
            limpar_lote_casos()
            historico['proximo_numero_por_analista'] = {}
            salvar_historico(historico)
            st.rerun()

    st.markdown("---")
    with st.expander("🔗 Combinar lote com outro analista"):
        st.caption(
            "Troque o arquivo exportado só por um canal já confiável da empresa "
            "(ex.: Teams/e-mail corporativo)."
        )

        st.markdown("**Exportar meu lote**")
        if not lote_atual:
            st.caption("Seu lote está vazio — nada para exportar ainda.")
        else:
            st.download_button(
                "⬇️ Baixar lote para enviar",
                data=exportar_lote_arquivo(lote_atual),
                file_name=f"lote_coaf_export_{datetime.now().strftime('%Y%m%d_%H%M')}.coafbatch",
                mime="application/octet-stream",
                key="btn_exportar_lote",
            )

        st.markdown("---")
        st.markdown("**Importar e mesclar lote de outra pessoa**")
        arquivo_lote_externo = st.file_uploader(
            "Arquivo .coafbatch recebido", type=["coafbatch"], key="upload_lote_externo"
        )

        if arquivo_lote_externo and st.button("🔀 Mesclar no meu lote", key="btn_merge_lote"):
            try:
                casos_externos = importar_lote_arquivo(arquivo_lote_externo.read())
            except (json.JSONDecodeError, UnicodeDecodeError):
                st.error("Arquivo corrompido ou inválido.", icon=":material/error:")
            else:
                mesclado, colisoes = mesclar_lotes_bloqueando_colisao(lote_atual, casos_externos)
                if colisoes:
                    st.error(
                        f"{len(colisoes)} número(s) de caso aparecem nos dois lados — "
                        "mesclagem BLOQUEADA para não perder nenhum dos dois. Renumere um "
                        "dos lados (faixa do analista errado?) e tente de novo.",
                        icon=":material/error:",
                    )
                    st.table(colisoes)
                else:
                    salvar_lote_casos(mesclado)
                    st.success(
                        f"{len(casos_externos)} caso(s) mesclado(s). "
                        f"Lote agora tem {len(mesclado)} caso(s) no total.",
                        icon=":material/check_circle:",
                    )
                    st.rerun()

# ========================
# CONSULTAR REPORTES JÁ ARQUIVADOS (substitui procurar no documento mestre)
# ========================
with st.expander("📚 Consultar reportes já enviados (arquivo permanente)"):
    arquivo_casos = carregar_arquivo_casos()
    if not arquivo_casos:
        st.caption("Nenhum reporte arquivado ainda — toda vez que um caso é adicionado ao "
                   "lote XML (pelo formulário, manualmente, ou por importação), uma cópia "
                   "completa fica guardada aqui para sempre, mesmo depois do lote ser limpo.")
    else:
        def _analista_do_numero_caso(numero_caso):
            for nome_analista, inicio in FAIXA_NUMERO_POR_ANALISTA.items():
                if inicio < int(numero_caso or 0) <= inicio + TAMANHO_FAIXA_POR_ANALISTA:
                    return nome_analista
            return "?"

        col_busca1, col_busca2, col_busca3, col_busca4 = st.columns(4)
        with col_busca1:
            filtro_texto = st.text_input("Buscar por nome ou CPF", key="filtro_arquivo_texto")
        with col_busca2:
            filtro_periodo = st.text_input("Buscar por período (ex: 06/2026)", key="filtro_arquivo_periodo")
        with col_busca3:
            filtro_analista = st.selectbox(
                "Analista", ["Todos"] + list(FAIXA_NUMERO_POR_ANALISTA.keys()), key="filtro_arquivo_analista"
            )
        with col_busca4:
            filtro_status = st.selectbox("Status", ["Todos", "Rascunho", "Enviado"], key="filtro_arquivo_status")

        algum_filtro_ativo = bool(filtro_texto) or bool(filtro_periodo) or filtro_analista != "Todos" or filtro_status != "Todos"

        resultados = []
        if algum_filtro_ativo:
            resultados = arquivo_casos
            if filtro_texto:
                termo = filtro_texto.strip().lower()
                termo_cpf = re.sub(r'\D', '', filtro_texto)
                resultados = [
                    c for c in resultados
                    if termo in c.get("nome", "").lower()
                    or (termo_cpf and termo_cpf in re.sub(r'\D', '', c.get("cpf", "")))
                ]
            if filtro_periodo:
                resultados = [
                    c for c in resultados
                    if filtro_periodo.strip() in f"{c.get('dt_inicio', '')} {c.get('dt_fim', '')}"
                ]
            if filtro_analista != "Todos":
                resultados = [c for c in resultados if _analista_do_numero_caso(c.get("numero_caso")) == filtro_analista]
            if filtro_status != "Todos":
                status_alvo = "rascunho" if filtro_status == "Rascunho" else "enviado"
                resultados = [c for c in resultados if c.get("status", "rascunho") == status_alvo]
            resultados = sorted(resultados, key=lambda c: c.get("arquivado_em", ""), reverse=True)
            st.caption(f"{len(resultados)} de {len(arquivo_casos)} reporte(s) arquivado(s) no total "
                       f"(mostrando até 50).")
        else:
            st.caption(f"{len(arquivo_casos)} reporte(s) arquivado(s) no total. "
                       "Preencha algum filtro acima para listar.")

        for caso in resultados[:50]:
            chave_widget = f"{caso.get('numero_caso')}_{caso.get('dt_inicio')}_{caso.get('dt_fim')}"
            status_caso = caso.get("status", "rascunho")
            selo_status = "🔓 Rascunho" if status_caso == "rascunho" else "🔒 Enviado"
            titulo = (f"Caso {caso.get('numero_caso')} — {caso.get('nome', '')} "
                      f"({_analista_do_numero_caso(caso.get('numero_caso'))}) — {selo_status}")
            with st.expander(titulo):
                st.caption(f"Período: {caso.get('dt_inicio', '')} a {caso.get('dt_fim', '')} · "
                           f"CPF: {formatar_cpf(caso.get('cpf', ''))} · "
                           f"Arquivado em: {caso.get('arquivado_em', '-')}"
                           + (f" · Enviado em: {caso['enviado_em']}" if caso.get("enviado_em") else ""))

                if status_caso == "enviado":
                    st.text_area("Texto completo (travado — já enviado, não pode ser alterado)",
                                 caso.get("det", ""), height=200, disabled=True,
                                 key=f"arquivo_texto_{chave_widget}")
                else:
                    st.caption(":material/lock_open: Ainda em rascunho — pode editar e salvar, ou excluir, antes do envio do lote.")
                    texto_editado = st.text_area("Texto completo (editável)", caso.get("det", ""),
                                                  height=200, key=f"arquivo_texto_{chave_widget}")

                    col_editar, col_excluir = st.columns(2)
                    with col_editar:
                        if st.button("💾 Salvar edição", key=f"arquivo_salvar_{chave_widget}"):
                            editar_caso_arquivado(caso.get("dt_inicio"), caso.get("dt_fim"),
                                                   caso.get("numero_caso"), texto_editado)
                            st.success("Alterações salvas.")
                            st.rerun()
                    with col_excluir:
                        confirmar_exclusao = st.checkbox("Confirmar exclusão", key=f"arquivo_confirma_{chave_widget}")
                        if st.button("🗑️ Excluir caso", key=f"arquivo_excluir_{chave_widget}",
                                     disabled=not confirmar_exclusao):
                            excluir_caso_arquivado(caso.get("dt_inicio"), caso.get("dt_fim"), caso.get("numero_caso"))
                            st.success("Caso excluído.")
                            st.rerun()

                if DOCX_DISPONIVEL:
                    st.download_button(
                        "⬇️ Baixar este reporte em .docx",
                        data=gerar_docx_bytes(caso.get("det", ""), fonte="Montserrat", tamanho_pt=11),
                        file_name=f"reporte_coaf_{caso.get('numero_caso')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"arquivo_docx_{chave_widget}",
                    )

# ========================
# STEP 1: UPLOAD PDFs
# ========================
st.subheader("📂 1. Carregar PDFs do time de risco")
st.caption("Você pode carregar múltiplos PDFs (uma ou mais marcas) — os valores serão somados automaticamente.")
pdf_files = st.file_uploader("Faça upload dos PDFs", type=["pdf"], accept_multiple_files=True)

# ========================
# STEP 2: BUSCAR CPF
# ========================
st.subheader("🔍 2. Buscar usuário pelo CPF")
cpf_busca = st.text_input("Digite o CPF", placeholder="000.000.000-00", key="cpf_busca_input")

dados_consolidados = {}
detalhes_por_casa = []

if pdf_files and cpf_busca and len(re.sub(r'\D', '', cpf_busca)) >= 11:
    with st.spinner(f"Buscando CPF em {len(pdf_files)} PDF(s)..."):
        detalhes_por_casa = buscar_cpf_em_pdfs(pdf_files, cpf_busca)

        if detalhes_por_casa:
            nome_final = detalhes_por_casa[0]['nome']
            cpf_final = detalhes_por_casa[0]['cpf']

            total_dep = sum(d['total_dep'] for d in detalhes_por_casa)
            qtd_dep = sum(d['qtd_dep'] for d in detalhes_por_casa)
            total_saq = sum(d['total_saq'] for d in detalhes_por_casa)
            qtd_saq = sum(d['qtd_saq'] for d in detalhes_por_casa)
            val_cas = sum(d['val_cas'] for d in detalhes_por_casa)
            qtd_cas = sum(d['qtd_cas'] for d in detalhes_por_casa)
            val_sb = sum(d['val_sb'] for d in detalhes_por_casa)
            qtd_sb = sum(d['qtd_sb'] for d in detalhes_por_casa)

            dados_consolidados = {
                'nome': nome_final,
                'cpf': cpf_final,
                'total_dep': total_dep,
                'qtd_dep': qtd_dep,
                'total_saq': abs(total_saq),
                'qtd_saq': qtd_saq,
                'val_cas': val_cas,
                'qtd_cas': qtd_cas,
                'val_sb': val_sb,
                'qtd_sb': qtd_sb,
            }

            st.success(f"Encontrado em {len(detalhes_por_casa)} casa(s): **{nome_final}**", icon=":material/check_circle:")

            with st.expander("📊 Ver detalhes por casa de apostas"):
                for d in detalhes_por_casa:
                    st.markdown(f"**{d['casa']}**")
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Depósitos", f"{formatar_valor(d['total_dep'])} ({d['qtd_dep']}x)")
                    with col2:
                        st.metric("Saques", f"{formatar_valor(d['total_saq'])} ({d['qtd_saq']}x)")
                    with col3:
                        sinal_cas = "📈 Ganho" if d['val_cas'] >= 0 else "📉 Perda"
                        st.metric("Cassino", f"{formatar_valor(d['val_cas'])} ({sinal_cas})")
                    with col4:
                        sinal_sb = "📈 Ganho" if d['val_sb'] >= 0 else "📉 Perda"
                        st.metric("Esportivas", f"{formatar_valor(d['val_sb'])} ({sinal_sb})")
                    st.markdown("---")

                st.markdown("**📊 TOTAL CONSOLIDADO**")
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total depósitos", f"{formatar_valor(total_dep)} ({qtd_dep}x)")
                with col2:
                    st.metric("Total saques", f"{formatar_valor(total_saq)} ({qtd_saq}x)")
                with col3:
                    sinal = "📈 Ganho" if val_cas >= 0 else "📉 Perda"
                    st.metric("Cassino líquido", f"{formatar_valor(val_cas)} ({sinal})")
                with col4:
                    sinal = "📈 Ganho" if val_sb >= 0 else "📉 Perda"
                    st.metric("Esportivas líquido", f"{formatar_valor(val_sb)} ({sinal})")
        else:
            st.warning("CPF não encontrado em nenhum PDF. Preencha os campos manualmente.", icon=":material/warning:")

# ========================
# STEP 3: FORMULÁRIO
# ========================
st.markdown("---")
st.subheader("📝 3. Preencher dados do caso")

col1, col2 = st.columns(2)

with col1:
    st.markdown("**Identificação**")
    analista = st.selectbox("Analista", ["Analista A", "Analista B", "Outro"])
    numero_caso = st.number_input(
        "Nº do caso", min_value=1, step=1,
        value=proximo_numero_para_analista(historico, analista),
        key=f"numero_caso_input_{analista}",
        help="Cada analista tem uma faixa própria de números (ver barra lateral) para nunca "
             "colidir com o número usado por outra pessoa ao mesclar os lotes.",
    )
    if not numero_esta_na_faixa_do_analista(int(numero_caso), analista):
        _faixa_ini = FAIXA_NUMERO_POR_ANALISTA.get(analista, FAIXA_NUMERO_POR_ANALISTA["Outro"])
        st.caption(
            f":material/warning: Fora da faixa reservada para {analista} "
            f"({_faixa_ini + 1}–{_faixa_ini + TAMANHO_FAIXA_POR_ANALISTA}). Confirme se é "
            "intencional (ex.: caso anterior à nova numeração por faixas)."
        )
    periodo = st.text_input("Período de referência", value="MÊS 06/2026")
    nome = st.text_input("Nome completo", value=dados_consolidados.get('nome', ''))
    genero_opcao = st.selectbox(
        "Gênero para redação",
        ["Automático", "Masculino (o apostador)", "Feminino (a apostadora)"],
        key="genero_opcao_input",
        help="Define a flexão de gênero no texto formal ('o apostador' vs 'a apostadora'). O modo automático detecta pelo primeiro nome."
    )
    cpf_fmt = dados_consolidados.get('cpf', formatar_cpf(re.sub(r'\D', '', cpf_busca)) if cpf_busca else '')
    cpf = st.text_input("CPF", value=cpf_fmt)
    data_nasc = st.text_input("Data de nascimento", key="data_nasc_raw",
        placeholder="Digite só os números, ex: 15031990", on_change=_mask_data_nasc)
    cidade_uf = st.text_input("Cidade/UF", placeholder="Ex: SÃO PAULO/SP", key="cidade_uf_input")
    pep = st.selectbox("PEP", ["NÃO", "SIM"], key="pep_input")

with col2:
    st.markdown("**Movimentação financeira**")
    renda_str = st.text_input("Renda presumida (R$)", key="renda_raw",
        placeholder="Ex: 1621 (vira 1.621,00) ou 1621,50 (vira 1.621,50)", on_change=_mask_renda)

    col_d1, col_d2 = st.columns(2)
    with col_d1:
        qtd_dep = st.number_input("Qtd. depósitos", min_value=0, step=1,
            value=int(dados_consolidados.get('qtd_dep', 0)))
    with col_d2:
        dep_val = dados_consolidados.get('total_dep', 0)
        dep_str_default = f"{dep_val:,.2f}".replace(",","X").replace(".",",").replace("X",".") if dep_val else ""
        total_dep_str = st.text_input("Total depositado (R$)", value=dep_str_default)

    col_s1, col_s2 = st.columns(2)
    with col_s1:
        qtd_saq = st.number_input("Qtd. saques", min_value=0, step=1,
            value=int(dados_consolidados.get('qtd_saq', 0)))
    with col_s2:
        saq_val = dados_consolidados.get('total_saq', 0)
        saq_str_default = f"{abs(saq_val):,.2f}".replace(",","X").replace(".",",").replace("X",".") if saq_val else ""
        total_saq_str = st.text_input("Total sacado (R$)", value=saq_str_default)

st.markdown("---")
col3, col4 = st.columns(2)

with col3:
    st.markdown("**⚽ Apostas esportivas**")
    qtd_esp = st.number_input("Qtd. apostas esportivas", min_value=0, step=1,
        value=int(dados_consolidados.get('qtd_sb', 0)))

    val_sb = dados_consolidados.get('val_sb', 0)
    sb_str_default = f"{abs(val_sb):,.2f}".replace(",","X").replace(".",",").replace("X",".") if val_sb else ""
    res_esp_str = st.text_input("Resultado esportivo (R$)", value=sb_str_default)

    tipo_esp_auto = "ganhos" if val_sb > 0 else ("perdas" if val_sb < 0 else "perdas")
    tipo_esp_idx = 0 if tipo_esp_auto == "perdas" else (1 if tipo_esp_auto == "ganhos" else 2)
    tipo_esp = st.selectbox("Tipo resultado esportivo", ["perdas", "ganhos", "nenhum"], index=tipo_esp_idx)

with col4:
    st.markdown("**🎰 Jogos online / Cassino**")
    qtd_cas = st.number_input("Qtd. rodadas cassino", min_value=0, step=1,
        value=int(dados_consolidados.get('qtd_cas', 0)))
    sem_rodadas = st.checkbox("Sem registro de número de rodadas")

    val_cas = dados_consolidados.get('val_cas', 0)
    cas_str_default = f"{abs(val_cas):,.2f}".replace(",","X").replace(".",",").replace("X",".") if val_cas else ""
    res_cas_str = st.text_input("Resultado cassino (R$)", value=cas_str_default)

    tipo_cas_auto = "ganhos" if val_cas > 0 else ("perdas" if val_cas < 0 else "perdas")
    tipo_cas_idx = 0 if tipo_cas_auto == "perdas" else (1 if tipo_cas_auto == "ganhos" else 2)
    tipo_cas = st.selectbox("Tipo resultado cassino", ["perdas", "ganhos", "nenhum"], index=tipo_cas_idx)

st.markdown("---")
st.subheader("🔍 Background Check")

col5, col6 = st.columns(2)

with col5:
    tem_cnpj = st.checkbox("Possui CNPJ", key="tem_cnpj_input")
    empresas = []
    if tem_cnpj:
        qtd_empresas = st.number_input("Quantas empresas?", min_value=1, max_value=MAX_EMPRESAS, step=1,
            key="qtd_empresas_input")
        for i in range(int(qtd_empresas)):
            st.markdown(f"**Empresa {i+1}**")
            cnpj_emp_i = st.text_input(f"CNPJ #{i+1}", placeholder="00.000.000/0001-00", key=f"cnpj_emp_{i}")
            tipo_empresa_i = st.selectbox(f"Tipo de empresa #{i+1}", ["MEI", "Microempresa", "Empresa", "Outro"],
                key=f"tipo_empresa_{i}")
            fat_str_i = st.text_input(f"Faturamento presumido #{i+1} (R$)", placeholder="Ex: 88.466,40", key=f"fat_{i}")
            participacao_i = st.text_input(f"% de participação #{i+1}", placeholder="Ex: 100% ou 50%",
                key=f"participacao_{i}")
            ramo_empresa_i = st.text_input(f"Ramo da empresa #{i+1} (opcional)", key=f"ramo_{i}")
            empresa_inapta_i = st.checkbox(f"Empresa #{i+1} inapta", key=f"inapta_{i}")
            if cnpj_emp_i.strip():
                empresas.append({
                    'cnpj': cnpj_emp_i, 'tipo': tipo_empresa_i, 'fat': fat_str_i,
                    'participacao': participacao_i, 'ramo': ramo_empresa_i, 'inapta': empresa_inapta_i,
                })
            st.markdown("---")

with col6:
    antecedente = st.checkbox("Possui antecedentes criminais", key="antecedente_input")
    detalhe_crim = st.text_area("Detalhes do antecedente", placeholder="Ex: Réu em ação penal...",
        key="detalhe_crim_input") if antecedente else ""
    antecedente_crime_fin = st.checkbox(
        "Antecedente é especificamente crime financeiro (corrupção, fraude, lavagem)",
        key="antecedente_crime_fin_input",
        help="Aplica o CodEnq 1410 no XML do lote, diferente de antecedente criminal genérico."
    ) if antecedente else False

    reporte_anterior = st.checkbox("Possui reporte anterior", key="reporte_anterior_input")
    periodos_anteriores = []
    if reporte_anterior:
        qtd_periodos_ant = st.number_input("Quantos períodos anteriores comparar?", min_value=1,
            max_value=MAX_PERIODOS_ANTERIORES, step=1, key="qtd_periodos_ant_input")
        st.caption("Preencha em ordem cronológica (do mais antigo para o mais recente).")
        for i in range(int(qtd_periodos_ant)):
            st.markdown(f"**Período anterior {i+1}**")
            col_pa1, col_pa2 = st.columns(2)
            with col_pa1:
                mes_i = st.selectbox(f"Mês #{i+1}", list(MESES_NOME.keys()),
                    format_func=lambda m: MESES_NOME[m], key=f"mes_ant_{i}")
            with col_pa2:
                ano_i = st.number_input(f"Ano #{i+1}", min_value=2020, max_value=2030,
                    step=1, key=f"ano_ant_{i}")
            dep_i_str = st.text_input(f"Total depositado #{i+1} (R$)", key=f"dep_ant_{i}")
            saq_i_str = st.text_input(f"Total sacado #{i+1} (R$) — opcional", key=f"saq_ant_{i}")
            periodos_anteriores.append({
                'mes': mes_i, 'ano': ano_i, 'dep_str': dep_i_str, 'saq_str': saq_i_str,
            })
            st.markdown("---")

    bpc = st.checkbox("Beneficiário de programa social", key="bpc_input")
    detalhe_bpc = st.text_input("Detalhes do benefício", key="detalhe_bpc_input") if bpc else ""

    pep_relacionado = st.checkbox("PEP relacionado", key="pep_relacionado_input")
    detalhe_pep = st.text_area("Detalhes do PEP relacionado", key="detalhe_pep_input") if pep_relacionado else ""

    negativado = st.checkbox("Negativado", key="negativado_input")
    valor_neg_str = st.text_input("Valor negativado (R$)", key="valor_neg_input") if negativado else ""

obs = st.text_area("Observações adicionais (opcional)", key="obs_input")

st.subheader("⚖️ Classificação da suspeita")
col_c1, col_c2, col_c3 = st.columns(3)
with col_c1:
    lavagem = st.checkbox("Lavagem de Dinheiro", value=True, key="lavagem_input")
with col_c2:
    financiamento = st.checkbox("Financiamento ao Terrorismo", key="financiamento_input")
with col_c3:
    manipulacao = st.checkbox("Manipulação de Resultados", key="manipulacao_input")

st.markdown("---")
st.subheader("📤 Google Sheets (opcional)")
enviar_sheets = st.checkbox("Enviar automaticamente este caso para o Google Sheets ao gerar o reporte", value=False)
if enviar_sheets:
    sheet_id = st.text_input("ID da planilha (reportes_coaf_junho)",
        help="Trecho da URL entre '/d/' e '/edit'")
    aba_sheets = st.text_input("Nome da aba", value="BASE_REPORTES")
    if not os.path.exists(OAUTH_CLIENT_FILE):
        st.warning(f"Arquivo de credenciais OAuth não encontrado: {OAUTH_CLIENT_FILE}. "
                   "Siga o passo a passo de configuração antes de habilitar o envio.",
                   icon=":material/warning:")
else:
    sheet_id = ""
    aba_sheets = "BASE_REPORTES"

# ========================
# GERAR REPORTE
# ========================

col_btn1, col_btn2 = st.columns([3, 1])
with col_btn1:
    gerar_clicado = st.button("🚀 Gerar Reporte", use_container_width=True)
with col_btn2:
    st.button("🧹 Limpar campos", use_container_width=True, on_click=limpar_campos_do_formulario,
        help="Limpa os dados manuais preenchidos (renda, data de nascimento, antecedentes, etc.) "
             "para o próximo caso, sem remover os PDFs já carregados.")

if gerar_clicado:
    renda = parse_valor(renda_str)
    total_dep = parse_valor(total_dep_str)
    total_saq = parse_valor(total_saq_str)
    res_esp = parse_valor(res_esp_str)
    res_cas = parse_valor(res_cas_str)
    valor_neg = parse_valor(valor_neg_str)

    mult_dep = calcular_multiplo(total_dep, renda)
    mult_saq = calcular_multiplo(total_saq, renda)
    ticket = calcular_ticket_medio(total_dep, qtd_dep)
    ticket_saq = calcular_ticket_medio(total_saq, qtd_saq)
    ticket_esp = calcular_ticket_medio(res_esp, qtd_esp)

    perda_esp = res_esp if tipo_esp == "perdas" else (-res_esp if tipo_esp == "ganhos" else 0)
    perda_cas = res_cas if tipo_cas == "perdas" else (-res_cas if tipo_cas == "ganhos" else 0)
    perda_total = perda_esp + perda_cas
    pct_perda = calcular_percentual(perda_total, total_dep) if perda_total > 0 else 0

    NOMES_FEM = {
        "MARIA","ANA","JULIA","JULIANA","JESSICA","PATRICIA","FERNANDA","AMANDA","GABRIELA","GIULIA",
        "ZAIRA","ACALILIA","KAREN","SABRINA","VIVIANE","EVELINE","GILDA","TACIANA","ROSANA","EMANUELLY",
        "FLAVIA","JAINE","DANIELA","CIDIANE","MARCIA","PALOMA","BRUNA","CAMILA","LETICIA","LARISSA",
        "NATALIA","ISABELA","BIANCA","GIOVANA","RAFAELA","RENATA","VANESSA","ALICE","HELENA","BEATRIZ",
        "CLAUDIA","CRISTIANE","LUCIANA","PRISCILA","SIMONE","TANIA","VERONICA",
        "INES","ISIS","JUCILENE","MARILENE","MARILEIDE","MARINETE","YASMIN","YASMIM","NOEMI","RUTH",
        "ESTER","RAQUEL","MIRIAM","SUELEN","ELAINE","ADRIANE","ADRIANA","LUANA","LUANE","IARA","ELIANE",
        "SOLANGE","REGINA","SANDRA","DENISE","MONICA","ROBERTA","GISELE","GISELLE","CAROLINE",
        "CAROLINA","THAIS","THAYS","TAYNARA","THAYNARA","YASMINE","LORENA","LORRANE","LAIS","LAYS",
        "ANDRESSA","ANDREZA","VITORIA","EDUARDA","LUIZA","LUISA","SOFIA","VALENTINA",
        "MANUELA","ALICIA","LARA","MELISSA","SARAH","SARA","LIVIA","NATHALIA","NATHALY",
    }
    NOMES_MASC_EXCECAO = {"LUCA"}

    if genero_opcao == "Masculino (o apostador)":
        eh_feminino = False
    elif genero_opcao == "Feminino (a apostadora)":
        eh_feminino = True
    else:
        primeiro_nome_raw = nome.strip().split()[0].upper() if nome else ""
        _acentos = str.maketrans("ÁÀÂÃÉÊÍÓÔÕÚÇ", "AAAAEEIOOOUC")
        primeiro_nome = primeiro_nome_raw.translate(_acentos)

        if primeiro_nome in NOMES_FEM:
            eh_feminino = True
        elif primeiro_nome in NOMES_MASC_EXCECAO:
            eh_feminino = False
        elif len(primeiro_nome) > 2 and primeiro_nome.endswith("A"):
            eh_feminino = True
        else:
            eh_feminino = False

    genero = "a apostadora" if eh_feminino else "o apostador"
    genero_cap = "A apostadora" if eh_feminino else "O apostador"
    pron = "a mesma" if eh_feminino else "o mesmo"
    art = "da apostadora" if eh_feminino else "do apostador"

    # VALORES ENVOLVIDOS
    txt_valores = f"No período analisado foram registrados {qtd_dep} depósitos que totalizaram o montante de {formatar_valor(total_dep)}. "

    tem_esp = qtd_esp > 0 and tipo_esp != "nenhum"
    tem_cas = (qtd_cas > 0 or sem_rodadas) and tipo_cas != "nenhum"

    tipo_esp_txt = tipo_esp
    tipo_cas_txt = tipo_cas

    # BUG CORRIGIDO: antes, um .replace(",", ".") era aplicado no texto
    # inteiro (incluindo o valor monetário já formatado por formatar_valor,
    # que usa vírgula como separador decimal), transformando "82.984,07"
    # em "82.984.07" — a correção formata a quantidade de rodadas
    # ISOLADAMENTE, sem tocar no restante da string.
    qtd_cas_fmt = formatar_milhar(qtd_cas)

    if tem_esp and tem_cas:
        txt_valores += f"{genero_cap} efetuou {qtd_esp} apostas esportivas com {tipo_esp_txt} no valor de {formatar_valor(res_esp)} e registrou "
        if sem_rodadas:
            txt_valores += f"rodadas em jogos online com {tipo_cas_txt} no valor de {formatar_valor(res_cas)}. "
        else:
            txt_valores += f"{qtd_cas_fmt} rodadas em jogos online com {tipo_cas_txt} no valor de {formatar_valor(res_cas)}. "
    elif tem_esp:
        txt_valores += f"{genero_cap} efetuou {qtd_esp} apostas esportivas com {tipo_esp_txt} no valor de {formatar_valor(res_esp)} e não registrou movimentação em jogos online. "
    elif tem_cas:
        txt_valores += f"{genero_cap} não efetuou apostas esportivas e registrou "
        if sem_rodadas:
            txt_valores += f"rodadas em jogos online com {tipo_cas_txt} no valor de {formatar_valor(res_cas)}. "
        else:
            txt_valores += f"{qtd_cas_fmt} rodadas em jogos online com {tipo_cas_txt} no valor de {formatar_valor(res_cas)}. "
    else:
        txt_valores += f"{genero_cap} não efetuou apostas esportivas e não registrou movimentação em jogos online. "

    if qtd_saq > 0:
        txt_valores += f"Foram contabilizados {qtd_saq} saques no período, que somaram {formatar_valor(total_saq)}."

    # ANÁLISE INTERNA
    txt_analise = f"Ao realizar o Background Check {art}, foi identificado que {pron} "

    def descrever_empresa(emp):
        fat_val = parse_valor(emp['fat'])
        if emp['tipo'] == "MEI":
            desc = f"CNPJ {emp['cnpj']} (MEI)"
        else:
            ramo_txt = f" do ramo de {emp['ramo']}" if emp['ramo'] else ""
            desc = f"CNPJ {emp['cnpj']}{ramo_txt}"
        if fat_val > 0:
            desc += f", com faturamento presumido de {formatar_valor(fat_val)}"
        if emp['participacao']:
            desc += f" e participação de {emp['participacao']}"
        if emp['inapta']:
            desc += ", atualmente inapta"
        return desc

    if tem_cnpj and empresas:
        if len(empresas) == 1:
            emp = empresas[0]
            if emp['tipo'] == "MEI":
                txt_analise += f"é MEI inscrito no CNPJ sob o n° {emp['cnpj']}"
            else:
                ramo_txt = f" do ramo de {emp['ramo']}" if emp['ramo'] else ""
                txt_analise += f"é titular de {emp['tipo'].lower()}{ramo_txt} inscrita no CNPJ {emp['cnpj']}"
            fat_val = parse_valor(emp['fat'])
            if fat_val > 0:
                txt_analise += f", com faturamento presumido de {formatar_valor(fat_val)}"
                if emp['participacao']:
                    txt_analise += f" e participação de {emp['participacao']}"
            if emp['inapta']:
                txt_analise += ", atualmente inapta"
            txt_analise += f", possui renda pessoal presumida de {formatar_valor(renda)}"
        else:
            descricoes = "; ".join(descrever_empresa(e) for e in empresas)
            txt_analise += (f"é titular/sócio de {len(empresas)} pessoas jurídicas: {descricoes}, "
                            f"possui renda pessoal presumida de {formatar_valor(renda)}")
    else:
        txt_analise += f"não possui CNPJ registrado ou relacionado ao seu nome, possui renda pessoal presumida de {formatar_valor(renda)}"

    if bpc and detalhe_bpc:
        txt_analise += f". Destaca-se que {pron} foi beneficiário de programa de assistência social ({detalhe_bpc})"

    if pep == "SIM":
        txt_analise += " e é Pessoa Politicamente Exposta (PEP)"

    if antecedente and detalhe_crim:
        txt_analise += f". Destaca-se que {pron} {detalhe_crim}"
    elif antecedente:
        txt_analise += " e possui antecedentes criminais"

    if negativado and valor_neg > 0:
        txt_analise += f", com valores negativados junto a instituições financeiras no montante de {formatar_valor(valor_neg)}"

    if not antecedente:
        txt_analise += " e não responde nem respondeu a processos de natureza criminal"

    txt_analise += "."

    if pep_relacionado and detalhe_pep:
        txt_analise += f" Destaca-se ainda que foi identificada relação de parentesco ou vínculo com Pessoa Politicamente Exposta: {detalhe_pep}. Tal condição, nos termos da regulamentação vigente, eleva o grau de risco associado às operações identificadas."

    if obs:
        txt_analise += f" {obs}."

    # FUNDAMENTAÇÃO
    fracionamento_detectado = qtd_dep > 100
    diff_pct_paridade = None
    if total_dep > 0 and total_saq > 0:
        diff_pct_paridade = abs(total_dep - total_saq) / total_dep * 100

    if reporte_anterior and len(periodos_anteriores) > 0:
        ordinais_txt = {1:"segunda", 2:"terceira", 3:"quarta", 4:"quinta", 5:"sexta", 6:"sétima", 7:"oitava"}
        vez_ord = ordinais_txt.get(len(periodos_anteriores), f"{len(periodos_anteriores)+1}ª")
        abertura_txt = f"Os registros analisados confirmam a reiteração e o agravamento da incompatibilidade entre os recursos movimentados e a capacidade financeira {art}, ora comunicado a este Conselho pela {vez_ord} vez."
    else:
        abertura_escolhida = random.choice(ABERTURAS_FUNDAMENTACAO)
        continuacao_escolhida = random.choice(CONTINUACOES_ABERTURA_FUNDAMENTACAO).format(art=art)
        abertura_txt = f"{abertura_escolhida} {continuacao_escolhida}"
        if FRASES_COMPLETAS_FUNDAMENTACAO and random.random() < (1 / (len(ABERTURAS_FUNDAMENTACAO) + 1)):
            abertura_txt = random.choice(FRASES_COMPLETAS_FUNDAMENTACAO).format(art=art)

    caso_pouca_informacao = (not tem_cnpj and pep != "SIM" and not pep_relacionado
                              and not antecedente and not bpc and not negativado)
    usar_modelo_compacto = (
        not (reporte_anterior and len(periodos_anteriores) > 0)
        and caso_pouca_informacao and fracionamento_detectado
        and 0 < pct_perda < 40 and random.random() < 0.5
    )

    if usar_modelo_compacto:
        if pct_perda < 30:
            frase_perda = (f"as perdas registradas equivalem a apenas {formatar_pct(pct_perda)}% dos depósitos totais, "
                           f"o que reforça a hipótese de uso da plataforma como canal de circulação de recursos")
        else:
            frase_perda = (f"as perdas registradas equivalem a {formatar_pct(pct_perda)}% dos depósitos totais, percentual "
                           f"que, ainda que considerável, não afasta a incompatibilidade já evidenciada: independentemente "
                           f"do resultado das apostas, a origem dos valores aportados permanece sem respaldo compatível "
                           f"com a renda presumida {art}")
        txt_fund = (f"A análise dos dados consolidados revelou movimentações financeiras atípicas na conta {art}, "
                    f"incompatíveis com sua capacidade presumida e sem respaldo em fontes de renda ou empresas identificáveis. "
                    f"Chama atenção a alta frequência de depósitos, padrão indicativo de fracionamento de valores. "
                    f"Ademais, {frase_perda}. Na ausência de lastro financeiro legítimo, "
                    f"não se pode afastar a hipótese de que os recursos movimentados sejam de titularidade de terceiros ou "
                    f"de origem não justificável por fontes lícitas identificadas.")
    else:
        blocos = []

        if tem_cnpj:
            blocos.append(("Embora figure como titular de pessoa jurídica ativa, o faturamento empresarial não se confunde com sua renda pessoal disponível, tampouco indica, por si só, capacidade para a realização de aportes em valores elevados e recorrentes.", True))
        else:
            blocos.append(("Ausente qualquer fonte de renda identificável ou pessoa jurídica sob sua administração que justifique os aportes realizados, o volume transacionado não encontra respaldo em meios lícitos conhecidos.", True))

        if pep == "SIM":
            blocos.append(("A condição de Pessoa Politicamente Exposta eleva substancialmente o grau de risco associado às operações identificadas, nos termos da regulamentação vigente.", True))
        if pep_relacionado:
            blocos.append(("A condição de PEP relacionado eleva substancialmente o perfil de risco, reforçando a necessidade de comunicação a este Conselho.", True))
        if antecedente:
            blocos.append((f"A existência de antecedentes criminais identificados no Background Check agrava o perfil de risco {art}.", True))
        if fracionamento_detectado:
            blocos.append((f"Chama atenção, ainda, o elevado número de depósitos realizados no período ({qtd_dep}), padrão compatível com fracionamento de valores.", False))
        if diff_pct_paridade is not None and diff_pct_paridade <= 15:
            blocos.append(("Verifica-se, ainda, a quase paridade entre os valores depositados e sacados, de modo que os aportes realizados na plataforma são posteriormente sacados quase em sua integralidade.", False))

        random.shuffle(blocos)

        # BLOCO DE PERDAS — sempre por último, antes da conclusão.
        # AJUSTE (pendência resolvida): a partir de 30% de perda (faixas
        # 30-40% e >=40%), o framing passa a explicitar que, MESMO sendo um
        # percentual considerável, isso não é suficiente para justificar os
        # aportes frente à renda presumida — em vez de manter a mesma frase
        # de "circulação de recursos" usada para percentuais bem baixos.
        if pct_perda > 0:
            if pct_perda < 30:
                blocos.append((f"Destaca-se que as perdas totais no período correspondem a apenas {formatar_pct(pct_perda)}% do total depositado, comportamento atípico que reforça a hipótese de utilização da plataforma predominantemente como canal de circulação de recursos.", False))
            elif pct_perda < 40:
                blocos.append((f"Destaca-se que as perdas totais no período correspondem a {formatar_pct(pct_perda)}% do total depositado, percentual que, ainda que considerável, não afasta a incompatibilidade já evidenciada: independentemente do resultado das apostas, a origem dos valores aportados permanece sem respaldo compatível com a renda presumida {art}.", False))
            else:
                blocos.append((f"Destaca-se que as perdas totais no período correspondem a {formatar_pct(pct_perda)}% do total depositado, percentual expressivo que não afasta a incompatibilidade identificada: independentemente do resultado das apostas, a origem dos recursos aportados permanece sem justificativa compatível com a renda presumida {art}, evidenciando incompatibilidade entre a capacidade financeira presumida e o volume de recursos efetivamente consumido na plataforma.", False))
        elif tem_esp or tem_cas:
            blocos.append(("Destaca-se que não há perdas totais no período, o que nesse contexto pode caracterizar manutenção de valores ou baixa exposição ao risco.", False))

        CONECTORES_BLOCO = [None, "Além disso, ", "Some-se a isso, ", "Cumpre registrar que ", "Ademais, "]

        def _montar_blocos_com_conectores(lista_blocos):
            partes = []
            for i, (bloco, permite_conector) in enumerate(lista_blocos):
                if i == 0 or not permite_conector:
                    partes.append(bloco)
                    continue
                conector = random.choice(CONECTORES_BLOCO)
                if conector is None:
                    partes.append(bloco)
                else:
                    partes.append(conector + bloco[0].lower() + bloco[1:])
            return " ".join(partes)

        txt_fund = abertura_txt + " " + _montar_blocos_com_conectores(blocos)

        sem_empresa_de_peso = (not tem_cnpj) or all(e['tipo'] == "MEI" for e in empresas)
        if sem_empresa_de_peso:
            txt_fund += " " + random.choice(FECHAMENTOS_FUNDAMENTACAO)
        else:
            txt_fund += " " + FECHAMENTOS_FUNDAMENTACAO[0]

    # MÚLTIPLOS
    txt_mult = f"O valor total depositado no período analisado corresponde a mais de {mult_dep} vezes a renda mensal presumida {art}"
    if total_saq > 0:
        txt_mult += f", enquanto os saques totalizam aproximadamente {mult_saq} vezes essa renda"
    txt_mult += f". Além disso, a frequência e o volume das operações, com {qtd_dep} depósitos apresentando ticket médio de {formatar_valor(ticket)} em um período de 30 dias, podem indicar indícios de utilização da plataforma para finalidades diversas do simples entretenimento."

    # COMPARAÇÃO COM PERÍODOS ANTERIORES (múltiplos períodos, com mês/ano)
    # Monta a comparação cronológica entre todos os períodos anteriores
    # informados e o período atual — no padrão usado manualmente nos
    # reportes 38 e 65 do caso MARCOS VENICIO (ex: "corresponderam a 40
    # vezes a renda presumida em fevereiro/2026 e a 52 vezes em
    # abril/2026, recuaram para 35 vezes no período ora analisado").
    NUMERO_EXTENSO = {2: "dois", 3: "três", 4: "quatro", 5: "cinco", 6: "seis"}

    def _numero_periodos_extenso(n):
        return NUMERO_EXTENSO.get(n, str(n))

    def _montar_pontos_anteriores(pontos):
        """pontos: lista de (periodo_str, multiplo), SEM o período atual.
        Ex.: '40 vezes a renda presumida em fevereiro/2026 e 52 vezes em abril/2026'"""
        itens = []
        for i, (periodo, mult) in enumerate(pontos):
            if i == 0:
                itens.append(f"{mult} vezes a renda presumida em {periodo}")
            else:
                itens.append(f"{mult} vezes em {periodo}")
        if len(itens) == 1:
            return itens[0]
        return ", ".join(itens[:-1]) + " e " + itens[-1]

    if reporte_anterior and periodos_anteriores:
        pontos_dep_anteriores = []  # (periodo_extenso, multiplo) — sem o atual
        multiplos_saq_todos = []    # inclui o atual, na ordem cronológica

        for pa in periodos_anteriores:
            dep_pa = parse_valor(pa['dep_str'])
            saq_pa = parse_valor(pa['saq_str'])
            periodo_extenso = formatar_periodo_extenso(pa['mes'], pa['ano'])
            if dep_pa > 0:
                pontos_dep_anteriores.append((periodo_extenso, calcular_multiplo(dep_pa, renda)))
            if saq_pa > 0:
                multiplos_saq_todos.append(calcular_multiplo(saq_pa, renda))

        if total_saq > 0:
            multiplos_saq_todos.append(mult_saq)

        if pontos_dep_anteriores:
            tendencia = mult_dep - pontos_dep_anteriores[0][1]
            verbo = "recuaram" if tendencia < 0 else ("aumentaram" if tendencia > 0 else "mantiveram-se estáveis")
            txt_mult += (f" Em comparação aos períodos anteriormente reportados, verifica-se que os depósitos, "
                         f"que corresponderam a {_montar_pontos_anteriores(pontos_dep_anteriores)}, {verbo} "
                         f"para {mult_dep} vezes no período ora analisado.")

        if len(multiplos_saq_todos) > 1:
            lista_saq_txt = ", ".join(str(m) for m in multiplos_saq_todos[:-1]) + f" e {multiplos_saq_todos[-1]}"
            txt_mult += (f" Os saques, por sua vez, corresponderam a {lista_saq_txt} vezes a renda presumida "
                         f"nos {_numero_periodos_extenso(len(multiplos_saq_todos))} períodos, respectivamente.")

        if pontos_dep_anteriores or len(multiplos_saq_todos) > 1:
            txt_mult += (" Tal padrão evidencia a manutenção do comportamento atípico de movimentação financeira "
                         f"incompatível com a capacidade econômica {art} ao longo dos períodos já reportados a este Conselho.")

    # FUNDAMENTAÇÃO LEGAL
    itens_fund_legal = [
        "Incompatibilidade entre a renda estimada e os valores movimentados",
        "Aporte de valores sobre os quais recaia suspeita quanto à sua origem",
    ]
    if antecedente: itens_fund_legal.append("Existência de antecedente criminal")
    if reporte_anterior: itens_fund_legal.append("Reporte anterior com reiteração de conduta")
    if bpc: itens_fund_legal.append("Beneficiário de programa social com movimentação financeira incompatível")
    if pep == "SIM": itens_fund_legal.append("Condição de Pessoa Politicamente Exposta (PEP)")
    if pep_relacionado: itens_fund_legal.append("PEP relacionado com elevação do grau de risco da operação")
    if negativado: itens_fund_legal.append("Valores negativados junto a instituições financeiras")
    if fracionamento_detectado: itens_fund_legal.append("Fracionamento de depósitos")

    fund_legal = "A comunicação é realizada com base nas obrigações previstas na Portaria SPA/MF n. 1.143/24 e na Lei nº 9.613/1998 (art. 11) e nas normativas da Unidade de Inteligência Financeira (UIF/COAF), em razão de:\n"
    linhas_fund_legal = []
    for i, item in enumerate(itens_fund_legal):
        pontuacao = ";" if i == len(itens_fund_legal) - 1 else ""
        linhas_fund_legal.append(f"{item}{pontuacao}")
    fund_legal += "\n".join(linhas_fund_legal)

    cl = lambda x: "[X]" if x else "[ ]"

    reporte = f"""{int(numero_caso)}) Nome completo: {nome.upper()}
CPF: {cpf}
Data de nascimento: {data_nasc}
Cidade/UF: {cidade_uf.upper()}
PEP: {pep}

DESCRIÇÃO DA OPERAÇÃO E MOTIVAÇÃO DA COMUNICAÇÃO

Período da(s) operação(ões): {periodo}
Valor(es) envolvido(s): {txt_valores}

Análise interna: {txt_analise}

{txt_fund}

{txt_mult}

Tais características podem configurar indícios de lavagem de dinheiro ou uso indevido da plataforma para circulação de recursos não justificados.

Classificação da suspeita:
{cl(lavagem)} Lavagem de Dinheiro
{cl(financiamento)} Financiamento ao Terrorismo
{cl(manipulacao)} Manipulação de Resultados

Fundamentação legal: {fund_legal}"""

    # Guarda tudo que as etapas seguintes (decidir o resultado, baixar o
    # .docx, mesclar) vão precisar — em session_state, não em variáveis
    # locais. ARMADILHA CLÁSSICA DO STREAMLIT: `gerar_clicado` (um
    # st.button comum) só é True no exato rerun do clique; em QUALQUER
    # interação seguinte (marcar o rádio, clicar em outro botão) ele
    # volta a False e este bloco inteiro deixa de rodar. Sem guardar o
    # estado aqui, os botões abaixo pareciam "sumir" ou não fazer nada.
    st.session_state['reporte_atual'] = {
        'numero_caso': int(numero_caso),
        'analista': analista,
        'nome': nome,
        'cpf': cpf,
        'cidade_uf': cidade_uf,
        'periodo': periodo,
        'mult_dep': mult_dep,
        'mult_saq': mult_saq,
        'ticket': ticket,
        'ticket_saq': ticket_saq,
        'ticket_esp': ticket_esp,
        'pct_perda': pct_perda,
        'pep': pep,
        'pep_relacionado': pep_relacionado,
        'antecedente': antecedente,
        'antecedente_crime_fin': antecedente_crime_fin,
        'bpc': bpc,
        'fracionamento_detectado': fracionamento_detectado,
        'reporte_anterior': reporte_anterior,
        'periodos_anteriores': periodos_anteriores,
        'total_dep': total_dep,
        'total_saq': total_saq,
        'qtd_dep': qtd_dep,
        'qtd_saq': qtd_saq,
        'tem_cas': tem_cas,
        'qtd_cas': qtd_cas,
        'res_cas': res_cas,
        'tipo_cas': tipo_cas,
        'tem_esp': tem_esp,
        'qtd_esp': qtd_esp,
        'res_esp': res_esp,
        'tipo_esp': tipo_esp,
        'renda': renda,
        'lavagem': lavagem,
        'financiamento': financiamento,
        'manipulacao': manipulacao,
    }
    # Reseta os widgets abaixo para o novo caso — sem isso, gerar um
    # reporte novo continuaria mostrando o texto/escolha do caso anterior.
    st.session_state['texto_final_det_input'] = reporte
    st.session_state['resultado_final_radio'] = "Reportado ao COAF"
    st.session_state['confirma_resultado_final'] = False

# ========================
# DECIDIR O RESULTADO DO CASO (fora do "if gerar_clicado": precisa
# continuar visível nos reruns seguintes, ex. ao marcar o rádio/checkbox)
# ========================
if 'reporte_atual' in st.session_state:
    dados = st.session_state['reporte_atual']

    st.markdown("---")
    st.subheader("📊 Valores calculados")
    col_m1, col_m2, col_m3, col_m4 = st.columns(4)
    with col_m1: st.metric("Múltiplo depósitos", f"{dados['mult_dep']}x")
    with col_m2: st.metric("Múltiplo saques", f"{dados['mult_saq']}x")
    with col_m3: st.metric("Ticket médio dep.", formatar_valor(dados['ticket']))
    with col_m4: st.metric("% perdas / depósitos", f"{formatar_pct(dados['pct_perda'])}%")

    st.subheader("📄 Reporte gerado — revise e edite aqui se necessário")
    st.text_area(
        "Texto do reporte (editável — esta é a versão que vai para o .docx, a planilha e o lote XML)",
        height=500,
        key="texto_final_det_input",
    )
    st.success(f"Reporte {dados['numero_caso']} gerado.", icon=":material/check_circle:")

    if DOCX_DISPONIVEL:
        docx_bytes = gerar_docx_bytes(st.session_state["texto_final_det_input"], fonte="Montserrat", tamanho_pt=11)
        st.download_button(
            label="⬇️ Baixar reporte em .docx (Montserrat 11)",
            data=docx_bytes,
            file_name=f"reporte_coaf_{dados['numero_caso']}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            key="btn_baixar_docx_atual",
        )
        st.caption("Se o Montserrat não estiver instalado no computador que abrir o arquivo, "
                   "o Word substitui visualmente por outra fonte, mas o nome certo já fica salvo no documento.")
    else:
        st.info("Para gerar o .docx com a formatação já aplicada, instale a biblioteca: `py -m pip install python-docx`")

    st.markdown("---")
    st.subheader("📦 Resultado do caso")

    resultado_final = st.radio(
        "Resultado deste caso:", ["Reportado ao COAF", "Sem Indícios"],
        key="resultado_final_radio",
    )
    if resultado_final == "Reportado ao COAF":
        st.caption(
            "Vai para o lote XML do SISCOAF"
            + (" e para a planilha (Google Sheets), marcado como \"Reportado ao COAF\"." if enviar_sheets else ".")
        )
    else:
        st.caption(
            "NÃO entra no lote XML — este caso não será comunicado ao SISCOAF."
            + (" Mesmo assim vai para a planilha (Google Sheets), marcado como \"Sem Indícios\"." if enviar_sheets else "")
        )

    confirmar_resultado = st.checkbox(
        f'Confirmo que o resultado deste caso é: "{resultado_final}"',
        key="confirma_resultado_final",
    )

    def _montar_linha_sheets(dados, resultado_texto):
        resultado_cas_signed = -dados['res_cas'] if dados['tipo_cas'] == "perdas" else (dados['res_cas'] if dados['tipo_cas'] == "ganhos" else 0)
        resultado_esp_signed = -dados['res_esp'] if dados['tipo_esp'] == "perdas" else (dados['res_esp'] if dados['tipo_esp'] == "ganhos" else 0)
        principal_suspeita = (
            "Lavagem de Dinheiro" if dados['lavagem'] else
            "Financiamento ao Terrorismo" if dados['financiamento'] else
            "Manipulação de Resultados" if dados['manipulacao'] else ""
        )
        gatilhos_lista = ["Incompatibilidade Financeira"]
        if dados['pep'] == "SIM" or dados['pep_relacionado']: gatilhos_lista.append("PEP")
        if dados['antecedente']: gatilhos_lista.append("Antecedente Criminal")
        if dados['bpc']: gatilhos_lista.append("Beneficiário de Programas Sociais")
        if dados['manipulacao']: gatilhos_lista.append("Prêmio c/ Suspeita de Manipulação")
        return [
            dados['analista'],
            neutralizar_formula(dados['nome'].upper()),
            dados['cpf'],
            neutralizar_formula(dados['cidade_uf'].upper()),
            neutralizar_formula(dados['periodo']),
            date.today().strftime("%d/%m/%Y"),
            int(dados['qtd_dep']),
            dados['total_dep'],
            dados['ticket'],
            int(dados['qtd_saq']),
            dados['total_saq'],
            dados['ticket_saq'],
            "Sim" if dados['tem_cas'] else "Não",
            int(dados['qtd_cas']),
            resultado_cas_signed,
            "Sim" if dados['tem_esp'] else "Não",
            int(dados['qtd_esp']),
            resultado_esp_signed,
            dados['ticket_esp'],
            dados['renda'],
            dados['mult_dep'],
            principal_suspeita,
            ", ".join(gatilhos_lista),
            "SIM" if dados['reporte_anterior'] else "NÃO",
            len(dados['periodos_anteriores']),
            resultado_texto,
        ]

    def _enviar_para_sheets_se_marcado(resultado_texto):
        if not enviar_sheets:
            return
        if not sheet_id:
            st.error("Preencha o ID da planilha antes de gerar, ou desmarque o envio automático.",
                      icon=":material/error:")
            return
        linha = _montar_linha_sheets(dados, resultado_texto)
        with st.spinner("Enviando para o Google Sheets..."):
            ok, msg = enviar_para_sheets(sheet_id, aba_sheets, linha)
        if ok:
            st.success(f"Caso também registrado na planilha (aba {aba_sheets}) como \"{resultado_texto}\".",
                       icon=":material/check_circle:")
        else:
            st.error(f"Não foi possível enviar para o Sheets: {msg}", icon=":material/error:")

    def _registrar_no_historico(resultado_texto):
        historico['casos'].append({
            'numero': dados['numero_caso'],
            'analista': dados['analista'],
            'nome': dados['nome'].upper(),
            'cpf': dados['cpf'],
            'data': str(date.today()),
            'periodo': dados['periodo'],
            'resultado': resultado_texto,
        })
        historico.setdefault('proximo_numero_por_analista', {})[dados['analista']] = dados['numero_caso'] + 1
        salvar_historico(historico)

    if resultado_final == "Reportado ao COAF":
        if st.button("➕ Adicionar este caso ao lote XML", use_container_width=True,
                     disabled=not confirmar_resultado):
            cidade_partes = dados['cidade_uf'].upper().split("/")
            enquadramentos_caso = montar_enquadramentos(
                pep_flag=(dados['pep'] == "SIM"),
                fracionamento_flag=dados['fracionamento_detectado'],
                antecedente_crime_financeiro_flag=dados['antecedente_crime_fin'],
            )
            caso_para_lote = {
                "numero_caso": dados['numero_caso'],
                "cpf": re.sub(r'\D', '', dados['cpf']),
                "nome": sanitizar_texto_siscoaf(dados['nome'].upper()),
                "dt_inicio": periodo_texto_para_datas(dados['periodo'])[0].strftime("%d/%m/%Y"),
                "dt_fim": periodo_texto_para_datas(dados['periodo'])[1].strftime("%d/%m/%Y"),
                "ag_mun": sanitizar_texto_siscoaf(cidade_partes[0].strip()) if cidade_partes else "",
                "ag_uf": cidade_partes[1].strip() if len(cidade_partes) > 1 else "",
                # HIPÓTESE a confirmar: VlCred=prêmios(~saques) / VlDeb=apostas(~depósitos)
                "vl_cred": int(dados['total_saq']),
                "vl_deb": int(dados['total_dep']),
                "vl_prov": 0,
                "vl_prop": 0,
                "det": sanitizar_texto_siscoaf(st.session_state["texto_final_det_input"]),
                "enquadramentos": enquadramentos_caso,
                "pep_flag": 1 if dados['pep'] == "SIM" else 0,
            }
            adicionar_caso_ao_lote(caso_para_lote)
            _registrar_no_historico("Reportado ao COAF")
            _enviar_para_sheets_se_marcado("Reportado ao COAF")
            st.success(f"Caso {dados['numero_caso']} adicionado ao lote XML.")
            del st.session_state['reporte_atual']
            st.rerun()
    else:
        if st.button("✅ Registrar como Sem Indícios", use_container_width=True,
                     disabled=not confirmar_resultado):
            _registrar_no_historico("Sem Indícios")
            _enviar_para_sheets_se_marcado("Sem Indícios")
            st.success(f"Caso {dados['numero_caso']} registrado como Sem Indícios (não entra no lote XML).")
            del st.session_state['reporte_atual']
            st.rerun()

# ========================
# ADICIONAR CASO MANUAL AO LOTE (feito fora do gerador, ex: por outro analista)
# ========================
st.markdown("---")
with st.expander("✍️ Adicionar caso ao lote manualmente (reporte feito fora do gerador)"):
    st.caption("Use isso para casos já reportados por outra pessoa da equipe (ex: direto no "
               "Google Docs), que não passaram por este programa.")

    m_numero = st.number_input("Nº do caso", min_value=1, step=1, key="manual_numero")
    m_cpf = st.text_input("CPF", key="manual_cpf")
    m_nome = st.text_input("Nome completo", key="manual_nome")
    m_dt_inicio = st.date_input("Data inicial do período", key="manual_dt_inicio")
    m_dt_fim = st.date_input("Data final do período", key="manual_dt_fim")
    m_cidade_uf = st.text_input("Cidade/UF", placeholder="Ex: SÃO PAULO/SP", key="manual_cidade_uf")
    m_pep = st.selectbox("PEP", ["NÃO", "SIM"], key="manual_pep")
    m_vl_cred = st.number_input("VlCred (R$, inteiro)", min_value=0, step=1, key="manual_vl_cred")
    m_vl_deb = st.number_input("VlDeb (R$, inteiro)", min_value=0, step=1, key="manual_vl_deb")
    m_det = st.text_area("Texto da análise (Det) — cole o texto já usado no reporte", key="manual_det", height=200)
    m_fracionamento = st.checkbox("Fracionamento de depósitos", key="manual_fracionamento")
    m_antecedente_crime_fin = st.checkbox(
        "Antecedente criminal ligado a crime financeiro (corrupção, fraude, lavagem)",
        key="manual_antecedente_crime_fin",
        help="Marque quando o antecedente identificado for especificamente por crime "
             "econômico/financeiro (ex: corrupção ativa, fraude, lavagem em processo à "
             "parte) — aplica o CodEnq 1410, diferente de antecedente criminal genérico.")

    if st.button("➕ Adicionar caso manual ao lote", key="btn_add_manual"):
        cidade_partes = m_cidade_uf.upper().split("/")
        enquadramentos_manual = montar_enquadramentos(
            pep_flag=(m_pep == "SIM"),
            fracionamento_flag=m_fracionamento,
            antecedente_crime_financeiro_flag=m_antecedente_crime_fin,
        )

        caso_manual = {
            "numero_caso": int(m_numero),
            "cpf": re.sub(r'\D', '', m_cpf),
            "nome": sanitizar_texto_siscoaf(m_nome.upper()),
            "dt_inicio": m_dt_inicio.strftime("%d/%m/%Y"),
            "dt_fim": m_dt_fim.strftime("%d/%m/%Y"),
            "ag_mun": sanitizar_texto_siscoaf(cidade_partes[0].strip()) if cidade_partes else "",
            "ag_uf": cidade_partes[1].strip() if len(cidade_partes) > 1 else "",
            "vl_cred": int(m_vl_cred),
            "vl_deb": int(m_vl_deb),
            "vl_prov": 0,
            "vl_prop": 0,
            "det": sanitizar_texto_siscoaf(m_det),
            "enquadramentos": enquadramentos_manual,
            "pep_flag": 1 if m_pep == "SIM" else 0,
        }
        adicionar_caso_ao_lote(caso_manual)
        st.success(f"Caso {int(m_numero)} adicionado manualmente ao lote.")
        st.rerun()
