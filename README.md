# Gerador de Reportes COAF

Aplicação Streamlit que apoia o setor de PLD/FT de uma operadora de apostas de quota fixa na análise de apostadores suspeitos e na comunicação ao COAF. Consolida PDFs de risco por CPF, calcula métricas, gera a minuta formal do reporte (`.docx` e texto editável), registra o caso no Google Sheets e monta o lote XML no leiaute do SISCOAF.

Versão 1.0 — 01/10/2026. Ferramenta de apoio à análise; o enquadramento e o envio seguem sujeitos à validação de Compliance/Jurídico.

## Visão geral do fluxo

```
PDFs de risco (por CPF)
        |
        v
Consolidação + métricas (múltiplo sobre renda presumida, ticket médio, % de perdas)
        |
        v
Background check (formulário) --> texto do reporte (editável no app)
        |
        +--> .docx (Montserrat 11pt)
        |
        v
Resultado do caso
   |-- "Reportado ao COAF" --> lote_casos.json --> XML SISCOAF (ISO-8859-1)
   |                      \--> Google Sheets (opcional)
   \-- "Sem Indícios"     --> historico_casos.json (+ Google Sheets, opcional)
```

## Funcionamento

### Análise e geração do reporte

- Upload e consolidação de PDFs de risco por CPF.
- Cálculo de métricas: múltiplos sobre renda presumida, ticket médio, percentual de perdas.
- Texto formal gerado com pool variado de aberturas, conectores e fechamentos, editável dentro do app.
- A seleção do resultado ("Reportado ao COAF" ou "Sem Indícios") exige caixa de confirmação antes de habilitar o botão de ação.

### Lote XML SISCOAF

- Códigos de enquadramento (`CodEnq`) atribuídos conforme o texto e as métricas do caso: 1418, 1420, 1426, 1415 e 1410 (Portaria SPA/MF nº 1.143/2024).
- Toda entrada de texto é sanitizada antes de compor o XML: o SISCOAF rejeita lotes com `< > & " ' |` (erro 525).
- Codificação `iso-8859-1`. Caracteres fora do Latin-1 são removidos para não gerar entidades numéricas (`&#...;`).
- `NumOcorrencia` é renumerado do zero a cada exportação; o número interno do caso segue faixa fixa por analista.
- Fallback do local do fato: sede da comunicante, quando a cidade do apostador é indeterminada.

### Dois analistas, um lote

Cada máquina mantém seu próprio lote. A opção "Combinar lote com outro analista" exporta e importa um `.coafbatch`. Se houver `numero_caso` repetido, a mesclagem é bloqueada e os números conflitantes são exibidos.

### Arquivo permanente

`arquivo_casos.json` guarda o texto completo de todo caso já adicionado a um lote. Casos em `rascunho` podem ser editados ou excluídos; depois do envio confirmado passam a `enviado` e ficam travados.

## Arquivos

| Arquivo | Função |
|---|---|
| `gerador_coaf.py` | Aplicação principal (Streamlit) |
| `extrair_lote_coaf.py`, `verificar_gaps.py` | Scripts legados para o documento mestre `.docx` de ciclos anteriores |
| `diagnosticar_pdf.py` | Diagnóstico de leitura de PDFs de risco |
| `test_gerador_coaf.py` | Testes automatizados (`pytest`) |
| `iniciar_gerador_coaf.bat`, `reiniciar_gerador_coaf.bat`, `iniciar_gerador_coaf_oculto.vbs`, `criar_atalho_area_trabalho.ps1` | Inicialização no Windows |
| `requirements.txt`, `requirements-dev.txt` | Dependências de uso e de desenvolvimento |

## Instalação

Requer Windows e Python 3.14+ (inicializador `py`).

```cmd
py -m pip install -r requirements.txt
```

Para o Google Sheets, crie um cliente OAuth 2.0 (fluxo de usuário) e salve como `oauth_client.json` na raiz. O `token_google.json` é gerado no primeiro login. Nenhum dos dois é versionado.

## Como rodar

```cmd
py -m streamlit run gerador_coaf.py
```

Ou execute `iniciar_gerador_coaf.bat`, que também cria `.streamlit/config.toml` restringindo o servidor a `127.0.0.1`.

### Testes

```cmd
py -m pip install -r requirements-dev.txt
py -m pytest -v
```

Os testes usam dados fictícios e isolam os arquivos de dados (`COAF_DATA_DIR` e `monkeypatch`). Nenhum teste deve tocar os JSON reais.

## Segurança e dados pessoais

- `historico_casos.json`, `lote_casos.json`, `arquivo_casos.json` e `token_google.json` são gravados criptografados com DPAPI do Windows (`pywin32`), presos à conta que os gravou. Sem `pywin32`, são gravados em texto puro, com aviso em tela.
- O servidor fica restrito a `127.0.0.1`. Não exponha o app na rede.
- Campos de texto livre passam por `neutralizar_formula()` antes de ir ao Google Sheets, para evitar formula injection.
- Dados de apostadores (nome, CPF, renda, movimentações), PDFs, `.docx`, planilhas e credenciais estão no `.gitignore` e não pertencem a este repositório. Use apenas dados fictícios em testes e exemplos.
- Referências: LGPD (Lei 13.709/2018), Lei 9.613/1998, Lei Complementar 105/2001.

## Limitações conhecidas

- Somente Windows (DPAPI e scripts de inicialização).
- Os códigos de enquadramento e o mapeamento de valores (`VlCred`, `VlDeb`, `VlProv`, `VlProp`) são hipóteses de homologação.
- O `.coafbatch` não tem senha; trocar apenas por canal corporativo confiável.
- O fluxo de OAuth usa conta de usuário; não há suporte a Service Account.

## Antes de usar em produção

- Validar o XML gerado no ambiente de treinamento do SISCOAF antes de qualquer envio em produção.
- Confirmar com Compliance/Jurídico o enquadramento e a redação dos reportes.
- Confirmar com a TI que a pasta do projeto não está sincronizada pelo OneDrive.
