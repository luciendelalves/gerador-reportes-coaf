"""
Testes automatizados do gerador_coaf.py.

Rodar com:  py -m pytest -v

IMPORTANTE: nenhum teste aqui toca os arquivos reais do projeto
(historico_casos.json, lote_casos.json, arquivo_casos.json). Testes que
mexem em arquivo redirecionam os caminhos do módulo para uma pasta
temporária (via monkeypatch ou COAF_DATA_DIR) — isso existe porque,
durante o desenvolvimento desta versão, testes manuais chegaram a
poluir o historico_casos.json real com dados de teste por falta desse
cuidado. Não remova esse isolamento.
"""
import os
import sys
import tempfile

import pytest

sys.path.insert(0, os.path.dirname(__file__))
sys.argv = ["streamlit"]  # evita o aviso "streamlit run" ao importar fora do runtime

# Isola até a carga inicial do módulo (a linha `historico = carregar_historico()`
# roda no import, antes de qualquer fixture existir).
_TMP_IMPORT_DIR = tempfile.mkdtemp(prefix="coaf_test_import_")
os.environ["COAF_DATA_DIR"] = _TMP_IMPORT_DIR

import gerador_coaf as g  # noqa: E402


@pytest.fixture
def isolar_arquivos(tmp_path, monkeypatch):
    """Redireciona HISTORICO_FILE/LOTE_CASOS_FILE/ARQUIVO_CASOS_FILE do
    módulo já importado para dentro de tmp_path (pasta exclusiva deste
    teste, apagada automaticamente pelo pytest depois)."""
    monkeypatch.setattr(g, "HISTORICO_FILE", str(tmp_path / "historico_casos.json"))
    monkeypatch.setattr(g, "LOTE_CASOS_FILE", str(tmp_path / "lote_casos.json"))
    monkeypatch.setattr(g, "ARQUIVO_CASOS_FILE", str(tmp_path / "arquivo_casos.json"))
    yield tmp_path


def _caso(numero_caso, det="texto de teste", nome="TESTE"):
    return {
        "numero_caso": numero_caso,
        "cpf": "12345678909",
        "nome": nome,
        "dt_inicio": "01/09/2026",
        "dt_fim": "30/09/2026",
        "ag_mun": "CIDADE EXEMPLO",
        "ag_uf": "PE",
        "vl_cred": 100,
        "vl_deb": 200,
        "vl_prov": 0,
        "vl_prop": 0,
        "det": det,
        "enquadramentos": [1418],
        "pep_flag": 0,
    }


# ==========================================================
# Funções puras: parsing e formatação de valores
# ==========================================================
class TestParseEFormatacao:
    def test_parse_valor_formato_br(self):
        assert g.parse_valor("1.234,56") == 1234.56

    def test_parse_valor_vazio(self):
        assert g.parse_valor("") == 0.0

    def test_parse_valor_negativo(self):
        assert g.parse_valor("-500,00") == -500.0
        assert g.parse_valor("(500,00)") == -500.0

    def test_formatar_valor(self):
        assert g.formatar_valor(1234.5) == "R$ 1.234,50"

    def test_formatar_cpf(self):
        assert g.formatar_cpf("12345678909") == "123.456.789-09"

    def test_formatar_cpf_incompleto_nao_formata(self):
        assert g.formatar_cpf("123") == "123"

    def test_calcular_multiplo(self):
        assert g.calcular_multiplo(10000, 2000) == 5
        assert g.calcular_multiplo(10000, 0) == 0

    def test_calcular_percentual(self):
        assert g.calcular_percentual(30, 100) == 30.0


# ==========================================================
# Sanitização SISCOAF (Erro 525) e Formula Injection
# ==========================================================
class TestSanitizacaoSiscoaf:
    @pytest.mark.parametrize("bruto,proibido", [
        ("Empresa & Cia", "&"),
        ('Ele disse "oi"', '"'),
        ("D'Oeste", "'"),
        ("<script>", "<"),
        ("a > b", ">"),
        ("a | b", "|"),
    ])
    def test_remove_caracteres_proibidos(self, bruto, proibido):
        assert proibido not in g.sanitizar_texto_siscoaf(bruto)

    def test_garantir_iso_8859_1_remove_fora_do_latin1(self):
        resultado = g.garantir_iso_8859_1("teste ☑ emoji \U0001F389")
        resultado.encode("iso-8859-1")  # não pode levantar UnicodeEncodeError


class TestNeutralizarFormula:
    @pytest.mark.parametrize("valor", ["=cmd|calc", "+1234", "-500", "@SUM(A1)"])
    def test_prefixa_caracteres_de_formula(self, valor):
        assert g.neutralizar_formula(valor).startswith("'")

    @pytest.mark.parametrize("valor", ["JOAO DA SILVA", "SAO PAULO/SP", ""])
    def test_nao_mexe_em_texto_normal(self, valor):
        assert g.neutralizar_formula(valor) == valor


# ==========================================================
# Numeração por analista (faixas fixas, evita colisão entre máquinas)
# ==========================================================
class TestNumeracaoPorAnalista:
    def test_proximo_numero_comeca_no_inicio_da_faixa(self):
        historico_vazio = {"proximo_numero_por_analista": {}}
        assert g.proximo_numero_para_analista(historico_vazio, "Analista A") == 10001
        assert g.proximo_numero_para_analista(historico_vazio, "Analista B") == 20001

    def test_proximo_numero_respeita_contador_existente(self):
        historico = {"proximo_numero_por_analista": {"Analista A": 10047}}
        assert g.proximo_numero_para_analista(historico, "Analista A") == 10047

    def test_numero_dentro_da_propria_faixa(self):
        assert g.numero_esta_na_faixa_do_analista(10005, "Analista A") is True
        assert g.numero_esta_na_faixa_do_analista(20005, "Analista A") is False

    def test_faixas_configuradas_nunca_se_sobrepoem(self):
        faixas = sorted(g.FAIXA_NUMERO_POR_ANALISTA.values())
        for inicio_a, inicio_b in zip(faixas, faixas[1:]):
            assert inicio_b - inicio_a >= g.TAMANHO_FAIXA_POR_ANALISTA


# ==========================================================
# Mesclar lotes de dois analistas (bloqueio de colisão de numero_caso)
# ==========================================================
class TestMesclarLotes:
    def test_mescla_sem_colisao(self):
        locais = [_caso(10001, nome="A")]
        externos = [_caso(20001, nome="B")]
        mesclado, colisoes = g.mesclar_lotes_bloqueando_colisao(locais, externos)
        assert colisoes == []
        assert [c["numero_caso"] for c in mesclado] == [10001, 20001]

    def test_bloqueia_quando_ha_numero_repetido(self):
        locais = [_caso(10001, nome="A")]
        externos = [_caso(10001, nome="B, numero errado")]
        mesclado, colisoes = g.mesclar_lotes_bloqueando_colisao(locais, externos)
        assert mesclado is None
        assert len(colisoes) == 1
        assert colisoes[0]["numero_caso"] == 10001


# ==========================================================
# Exportar / importar lote (.coafbatch, entre as duas máquinas)
# ==========================================================
class TestExportarImportarLote:
    def test_round_trip(self):
        casos = [_caso(10001, nome="TESTE")]
        exportado = g.exportar_lote_arquivo(casos)
        importado = g.importar_lote_arquivo(exportado)
        assert importado == casos


# ==========================================================
# XML do SISCOAF: NumOcorrencia é renumerado do zero, independente do
# numero_caso interno (faixa por analista)
# ==========================================================
class TestGerarXmlLote:
    def test_num_ocorrencia_renumerado_do_zero(self, isolar_arquivos):
        import xml.etree.ElementTree as ET

        g.salvar_lote_casos([_caso(20005), _caso(10002)])
        xml_bytes = g.gerar_xml_lote_bytes()
        root = ET.fromstring(xml_bytes)
        numeros = [oc.find("NumOcorrencia").text for oc in root.iter("OCORRENCIA")]
        assert numeros == ["1", "2"]

    def test_lote_vazio_retorna_none(self, isolar_arquivos):
        g.salvar_lote_casos([])
        assert g.gerar_xml_lote_bytes() is None


# ==========================================================
# Arquivo permanente: rascunho editável/excluível -> travado ao ser
# marcado como enviado (não pode divergir do que foi comunicado)
# ==========================================================
class TestArquivoPermanente:
    def test_novo_caso_entra_como_rascunho(self, isolar_arquivos):
        g.adicionar_caso_ao_lote(_caso(10001))
        assert g.carregar_arquivo_casos()[0]["status"] == "rascunho"

    def test_limpar_lote_preserva_arquivo_permanente(self, isolar_arquivos):
        g.adicionar_caso_ao_lote(_caso(10001))
        g.limpar_lote_casos()
        assert g.carregar_lote_casos() == []
        assert len(g.carregar_arquivo_casos()) == 1

    def test_editar_rascunho_funciona(self, isolar_arquivos):
        g.adicionar_caso_ao_lote(_caso(10001))
        g.editar_caso_arquivado("01/09/2026", "30/09/2026", 10001, "texto editado")
        assert g.carregar_arquivo_casos()[0]["det"] == "texto editado"

    def test_marcar_como_enviado_trava_edicao(self, isolar_arquivos):
        g.adicionar_caso_ao_lote(_caso(10001))
        g.marcar_casos_como_enviados(g.carregar_lote_casos())
        with pytest.raises(ValueError):
            g.editar_caso_arquivado("01/09/2026", "30/09/2026", 10001, "tentando editar")

    def test_marcar_como_enviado_trava_exclusao(self, isolar_arquivos):
        g.adicionar_caso_ao_lote(_caso(10001))
        g.marcar_casos_como_enviados(g.carregar_lote_casos())
        with pytest.raises(ValueError):
            g.excluir_caso_arquivado("01/09/2026", "30/09/2026", 10001)

    def test_excluir_rascunho_funciona(self, isolar_arquivos):
        g.adicionar_caso_ao_lote(_caso(10001))
        g.excluir_caso_arquivado("01/09/2026", "30/09/2026", 10001)
        assert g.carregar_arquivo_casos() == []

    def test_reprocessar_nao_sobrescreve_caso_ja_enviado(self, isolar_arquivos):
        g.adicionar_caso_ao_lote(_caso(10001, det="original"))
        g.marcar_casos_como_enviados(g.carregar_lote_casos())
        g.adicionar_caso_ao_lote(_caso(10001, det="tentativa de reabrir"))
        arquivo = g.carregar_arquivo_casos()
        assert len(arquivo) == 1
        assert arquivo[0]["det"] == "original"
        assert arquivo[0]["status"] == "enviado"


# ==========================================================
# Ponta a ponta via AppTest — os dois bugs reais já pegos nesta versão:
# 1) "Adicionar ao lote" não funcionava de verdade ao clicar (gerar_clicado
#    voltava a False em qualquer interação seguinte, sem session_state).
# 2) "Sem Indícios" não ia pra planilha com o rótulo certo.
# ==========================================================
@pytest.fixture
def isolar_apptest(tmp_path, monkeypatch):
    # COAF_DATA_DIR faz o AppTest (que executa o script do zero) gravar
    # em tmp_path; o monkeypatch.setattr faz as ASSERÇÕES (que usam o
    # módulo `g` já importado, com os caminhos fixados na importação)
    # lerem da mesma pasta — sem isso os dois lados apontam pra pastas
    # temporárias diferentes e o teste falha por isolamento, não por bug.
    monkeypatch.setenv("COAF_DATA_DIR", str(tmp_path))
    monkeypatch.setattr(g, "HISTORICO_FILE", str(tmp_path / "historico_casos.json"))
    monkeypatch.setattr(g, "LOTE_CASOS_FILE", str(tmp_path / "lote_casos.json"))
    monkeypatch.setattr(g, "ARQUIVO_CASOS_FILE", str(tmp_path / "arquivo_casos.json"))
    yield tmp_path


def _gerar_reporte(at, cpf, nome):
    at.text_input(key="cpf_busca_input").set_value(cpf)
    for ti in at.text_input:
        if ti.label == "Nome completo":
            ti.set_value(nome)
    for b in at.button:
        if "Gerar Reporte" in (b.label or ""):
            b.click()
    at.run(timeout=30)


class TestFluxoCompletoUI:
    def test_adicionar_ao_lote_funciona_ao_clicar(self, isolar_apptest):
        from streamlit.testing.v1 import AppTest

        at = AppTest.from_file("gerador_coaf.py")
        at.run(timeout=30)
        _gerar_reporte(at, "11111111111", "TESTE PYTEST UM")

        for c in at.checkbox:
            if c.key == "confirma_resultado_final":
                c.set_value(True)
        at.run(timeout=30)

        for b in at.button:
            if "Adicionar este caso ao lote XML" in (b.label or ""):
                b.click()
        at.run(timeout=30)

        assert not at.exception
        assert len(g.carregar_lote_casos()) == 1

    def test_sem_indicios_nao_entra_no_lote_mas_registra_historico(self, isolar_apptest):
        from streamlit.testing.v1 import AppTest

        at = AppTest.from_file("gerador_coaf.py")
        at.run(timeout=30)
        _gerar_reporte(at, "22222222222", "TESTE PYTEST DOIS")

        for r in at.radio:
            if r.key == "resultado_final_radio":
                r.set_value("Sem Indícios")
        at.run(timeout=30)

        for c in at.checkbox:
            if c.key == "confirma_resultado_final":
                c.set_value(True)
        at.run(timeout=30)

        for b in at.button:
            if "Registrar como Sem Ind" in (b.label or ""):
                b.click()
        at.run(timeout=30)

        assert not at.exception
        assert g.carregar_lote_casos() == []
        historico = g.carregar_historico()
        assert historico["casos"][-1]["resultado"] == "Sem Indícios"
