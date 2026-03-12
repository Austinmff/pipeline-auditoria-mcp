# Pipeline de Auditoria Fiscal — Mercado de Curto Prazo (MCP/CCEE)

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python)](https://python.org)
[![Pandas](https://img.shields.io/badge/Pandas-2.x-150458?logo=pandas)](https://pandas.pydata.org)
[![OpenPyXL](https://img.shields.io/badge/OpenPyXL-3.x-green)](https://openpyxl.readthedocs.io)
[![LGPD](https://img.shields.io/badge/LGPD-Dados%20Fictícios-orange)](https://www.planalto.gov.br/ccivil_03/_ato2015-2018/2018/lei/l13709.htm)

---

## ⚠️ Aviso de Privacidade e LGPD

> **Todos os dados presentes neste repositório são 100% fictícios**, gerados exclusivamente para fins de demonstração de portfólio.
>
> Nenhum dado real de agentes do mercado de energia, CNPJs, valores financeiros, resultados do MCP ou qualquer informação proveniente da CCEE, Secretaria da Fazenda ou qualquer órgão público foi utilizado, armazenado ou exposto neste projeto.
>
> A estrutura das planilhas (nomes de colunas e formato) foi reproduzida com base em relatórios de acesso público disponibilizados pela CCEE, conforme a **Lei de Acesso à Informação (Lei nº 12.527/2011)**. O tratamento e uso de dados pessoais segue os princípios da **Lei Geral de Proteção de Dados (Lei nº 13.709/2018 — LGPD)**.

---

## Contexto e Problema

Durante estágio no setor de fiscalização de energia de uma Secretaria da Fazenda estadual, identifiquei um gargalo operacional no processo de auditoria do Mercado de Curto Prazo (MCP):

Auditores precisavam, manualmente e agente por agente, acessar o sistema nacional da CCEE, extrair duas planilhas distintas — **CFZ004 (Consumo)** e **CFZ003 (Contabilização)** — cruzá-las e filtrar apenas os perfis com cargas no estado, para então extrair o **Resultado Final** financeiro proporcional àquelas cargas.

**O problema não era só a repetição manual.** O cruzamento ingênuo pelo CNPJ do agente retornaria o resultado financeiro total do agente no Brasil — não apenas a parcela atribuível às cargas estaduais. Isso geraria um número inflado em até 15x dependendo da concentração geográfica do agente.

---

## Objetivo

Automatizar o pipeline de preparação de dados para auditoria, entregando ao auditor o **Resultado Final proporcional** às cargas do estado — calculado mês a mês por perfil de agente — sem intervenção manual.

---

## A Solução Técnica

O desafio central é que a planilha de Contabilização não tem granularidade geográfica — o Resultado Final vem agregado por perfil, sem quebra por estado. A solução aplica rateio proporcional:

```
Resultado_Proporcional_UF = Resultado_Final × (Carga_UF_MWh / Carga_Total_MWh)
```

Calculado individualmente por **perfil de agente** e por **mês de competência**, garantindo que sazonalidades e variações mensais de carga sejam respeitadas.

---

## Arquitetura — ETL

```
CFZ004_Consumo.xlsx  ──┐
                        ├──► pipeline.py ──► relatorio_auditoria_SE_2022.xlsx
CFZ003_Contab.xlsx   ──┘
```

| Etapa | O que acontece |
|-------|---------------|
| **Extract** | Leitura das duas planilhas com tipagem explícita de CNPJs como string |
| **Transform** | Identificação dos perfis com carga no estado alvo → cruzamento por `CNPJ Agente + Perfil de Agente` → cálculo da proporção de carga por perfil/mês → aplicação sobre o Resultado Final |
| **Load** | Exportação em Excel com aba de detalhe mensal e aba de resumo consolidado por perfil |

---

## Por que o cruzamento usa CNPJ + Perfil de Agente?

Um agente pode ter múltiplos perfis (ex: `DISTRIBEX NORDESTE`, `DISTRIBEX NORTE`, `DISTRIBEX SUL`), cada um com resultado financeiro independente na contabilização. Usar apenas o CNPJ traria todos os perfis do agente, incluindo os sem carga no estado. O cruzamento por chave composta `(CNPJ, Perfil)` garante que apenas o perfil correto é selecionado.

---

## Estrutura do Projeto

```
portfolio-auditoria-mcp/
│
├── data/
│   ├── raw/
│   │   ├── CFZ004_Consumo_2022_SIMULADO.xlsx       # Planilha de consumo (dados fictícios)
│   │   └── CFZ003_Contabilizacao_2022_SIMULADO.xlsx # Planilha de contabilização (dados fictícios)
│   └── processed/
│
├── src/
│   ├── pipeline.py                  # Pipeline principal (ETL)
│   └── gerar_dados_simulados.py     # Gerador de dados fictícios para reprodução
│
├── output/
│   └── relatorio_auditoria_SE_2022.xlsx  # Relatório de saída para o auditor
│
├── requirements.txt
└── README.md
```

---

## Como Executar

```bash
pip install -r requirements.txt
```

**Gerar os dados simulados:**
```bash
python src/gerar_dados_simulados.py
```

**Executar o pipeline:**
```bash
python src/pipeline.py
```

O parâmetro `uf` no final do `pipeline.py` pode ser alterado para qualquer estado — o pipeline é agnóstico ao estado alvo.

---

## O que o Auditor Recebe

O relatório Excel contém duas abas:

**`Detalhe_SE`** — Registro mensal por perfil com: carga total nacional, carga em SE, proporção, Resultado Final nacional e Resultado Proporcional SE.

**`Resumo_por_Perfil`** — Totais anuais por perfil com linha de total consolidado. Resultados negativos destacados em vermelho.

---

## Impacto

| | Antes | Depois |
|--|-------|--------|
| Seleção dos agentes | Manual, empresa por empresa | Automático, filtro por UF na planilha de consumo |
| Cruzamento | Manual, sem garantia de precisão | Join por chave composta (CNPJ + Perfil) |
| Resultado financeiro entregue | Total nacional do agente (impreciso) | Proporcional às cargas estaduais (correto) |
| Rastreabilidade | Nenhuma | Log completo de cada etapa da execução |

---

## Tecnologias

- **Python 3.10+**
- **Pandas** — manipulação, agrupamento, merge e cálculo proporcional
- **OpenPyXL** — exportação do relatório Excel formatado
- **Logging** — rastreabilidade de cada etapa do pipeline

---

*Projeto desenvolvido por Austin | [LinkedIn](#) | [GitHub](#)*
