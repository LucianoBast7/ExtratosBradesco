# Extratos Bradesco – Automação de Coleta e Exportação de Dados Bancários

## 📌 Visão Geral

Este projeto automatiza a coleta de extratos bancários do Bradesco via Web, utilizando Selenium e reconhecimento de áudio para quebra de reCAPTCHA. Após a coleta, os saldos e movimentações são armazenados em banco de dados PostgreSQL e exportados em planilhas Excel.

---

## 🔧 Funcionalidades

- Login automático no portal Bradesco Empresas.
- Resolução de reCAPTCHA com reconhecimento de áudio (SpeechRecognition).
- Coleta de saldos e movimentações bancárias por CNPJ.
- Armazenamento dos dados em tabelas PostgreSQL.
- Geração de relatórios `.xlsx` e layout `.txt` compatível com o sistema Sinqia.

---

## ⚙️ Tecnologias Utilizadas

- **Python**
- **Selenium WebDriver**
- **SpeechRecognition**
- **BeautifulSoup**
- **SQLAlchemy + PostgreSQL**
- **pandas / openpyxl**

---

## 📁 Estrutura de Pastas

- `pasta_input`: mapa de carteiras com CNPJs
- `pasta_exportacao_extratos`: saída dos arquivos `.xlsx` e `.txt`
- `0. Log`: arquivos de log por execução e data

---

## 🔐 Variáveis Sensíveis

Centralizadas no arquivo:
```
variaveis/variaveis_estratos_bradesco.py
```
Incluem:
- Credenciais de acesso ao Bradesco
- Strings de conexão PostgreSQL
- Pastas de entrada e saída

---

## 🚀 Como Executar

1. Configure o arquivo de variáveis.
2. Execute o script principal:

```bash
python extratos_bradesco_V14.py
```

---

## 🗃️ Arquivos Gerados

- `ExtratosBradesco.xlsx`: saldos diários por fundo
- `MovimentacoesBradesco.xlsx`: detalhamento de movimentações
- `layout_sinqia_{data}.txt`: movimentações filtradas formatadas para Sinqia

---

## ⚠️ Observações

- A execução ocorre apenas em dias úteis (validação automática via calendário da Vortx).
- Fundos ignorados podem ser configurados via dicionário `empresas_ignoradas`.
- Logs detalhados são salvos por execução.

---
