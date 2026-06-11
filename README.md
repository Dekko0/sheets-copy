# 📋 Sistema de Cópia de Abas Excel

Aplicação web local (Streamlit) que **copia abas inteiras** de uma planilha
**modelo** para uma planilha **destino**, preservando **tudo**: fórmulas,
formatação, validações de dados, células mescladas, formatação condicional,
gráficos, comentários, hyperlinks, larguras/alturas, congelamento de painéis e
*named ranges*.

Pensada para o domínio de **eficiência energética** — transferir abas
padronizadas de um Relatório de Diagnóstico (`RD_MODELO.xlsx`) para planilhas de
cada projeto que seguem a mesma estrutura.

---

## ✨ Recursos

- Modelo embutido (`RD_MODELO.xlsx`) **ou** upload de um modelo customizado.
- Seleção de abas com **pré-seleção** das 7 abas padrão e **metadados por aba**
  (nº de fórmulas, validações, mescladas, formatação condicional, gráficos…).
- **Análise de dependências** antes de copiar: avisa quando uma aba referencia
  outra que não existe na destino, com exemplos de fórmulas.
- Dois **modos de cópia**: *Substituir completamente* (padrão) ou *Sobrescrever
  apenas as células do modelo*.
- Criação automática de abas inexistentes na destino.
- Preserva *named ranges* sem sobrescrever os já existentes.
- Download do arquivo final preservando a extensão (`.xlsx` / `.xlsm`).
- Os arquivos originais **nunca são modificados** (tudo em memória).

---

## 🧱 Stack

- Python 3.10+ (testado em **3.14.2**)
- [Streamlit](https://streamlit.io/) `1.52.2`
- [openpyxl](https://openpyxl.readthedocs.io/) `3.1.5`
- `pytest` + `pytest-cov` (testes)

> Não usamos `pandas` para a cópia (perde fórmulas/formatação), nem
> `xlwings`/`pywin32`/COM (a app roda *headless*, sem Excel instalado).

---

## 📁 Estrutura

```
.
├── app.py                      # Entrypoint Streamlit (UI em 5 passos)
├── requirements.txt            # Dependências de runtime (pinned)
├── requirements-dev.txt        # Dependências de teste
├── pyproject.toml              # Config do pytest/ruff/coverage
├── .streamlit/config.toml      # Tema e limite de upload
├── assets/
│   └── RD_MODELO.xlsx          # ⚠️ Modelo padrão (você adiciona — veja abaixo)
├── config/
│   └── defaults.py             # Modelo padrão, abas pré-selecionadas, constantes
├── core/
│   ├── exceptions.py           # Exceções do domínio
│   ├── workbook_loader.py      # Carrega de path/bytes; valida; keep_vba p/ .xlsm
│   ├── sheet_inspector.py      # Metadados das abas
│   ├── sheet_copier.py         # ❤️ Núcleo: cópia integral + relatório
│   └── dependency_checker.py   # Dependências entre abas
├── ui/
│   ├── components.py           # Componentes Streamlit reutilizáveis
│   └── styles.py               # CSS customizado
└── tests/                      # 46 testes (pytest), ~89% de cobertura
    ├── conftest.py             # Fixtures (workbooks gerados em memória)
    ├── fixtures/sample_template.xlsx
    ├── test_sheet_copier.py
    ├── test_dependency_checker.py
    ├── test_workbook_loader.py
    ├── test_sheet_inspector.py
    └── test_app_smoke.py
```

---

## 🚀 Instalação e execução

### 1. (Opcional, recomendado) Ambiente virtual

**Windows (PowerShell):**
```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
```

**Linux/macOS:**
```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 2. Instalar dependências
```bash
pip install -r requirements.txt
```

### 3. Adicionar o modelo padrão (opcional)
Coloque o arquivo **`RD_MODELO.xlsx`** dentro da pasta **`assets/`**.
Sem ele, a app continua funcionando — basta escolher *“Subir outro modelo”*.

### 4. Executar
```bash
streamlit run app.py
```

> No Windows, se `streamlit` não estiver no PATH: `py -m streamlit run app.py`.

A interface abre em `http://localhost:8501`.

---

## 🖱️ Como usar (fluxo em 5 passos)

```
┌─────────────────────────────────────────────────────────────┐
│  📋 Sistema de Cópia de Abas Excel                           │
├─────────────────────────────────────────────────────────────┤
│  PASSO 1 — Modelo                                            │
│   (●) Usar modelo padrão (RD_MODELO.xlsx)  ( ) Subir outro   │
│   ✅ Modelo ativo: RD_MODELO.xlsx — 15 abas · 320 KB         │
├─────────────────────────────────────────────────────────────┤
│  PASSO 2 — Planilha destino                                  │
│   [ arraste o .xlsx/.xlsm aqui ]                             │
│   ✅ Destino: relatorio_cliente_X.xlsx — 12 abas            │
├─────────────────────────────────────────────────────────────┤
│  PASSO 3 — Selecione as abas                                 │
│   ┌─ tabela: Aba | Fórmulas | Validações | Mescladas | ... ┐ │
│   [multiselect: 7 abas padrão já marcadas]                   │
│   🆕 Serão criadas na destino: FATURAS                       │
├─────────────────────────────────────────────────────────────┤
│  PASSO 4 — Dependências                                      │
│   ⚠️ 'UC_DADOS' referencia 'FATURAS' que não existe na       │
│      destino — ex.: UC_DADOS!B1 → =XLOOKUP(MAX(FATURAS!...   │
│   ☐ Entendo o risco e quero continuar mesmo assim           │
├─────────────────────────────────────────────────────────────┤
│   ⚙️ Modo de cópia: [Substituir completamente ▼]            │
│   [ ▶ Executar cópia ]                                       │
├─────────────────────────────────────────────────────────────┤
│  PASSO 5 — Resultado                                         │
│   ✅ 7 abas copiadas · 1 criada · 0 avisos                   │
│   [ 📥 Baixar planilha resultante ]                          │
└─────────────────────────────────────────────────────────────┘
```

### Modos de cópia

| Modo | Comportamento quando a aba já existe na destino |
|---|---|
| **Substituir completamente** (padrão) | A aba é recriada do zero (na mesma posição). Qualquer conteúdo antigo do destino é descartado. |
| **Sobrescrever apenas células do modelo** | A aba é mantida; as células do modelo são sobrescritas, mas células do destino **fora** do intervalo do modelo são preservadas. Mescladas/validações/formatação condicional são re-sincronizadas a partir do modelo. |

Em ambos os modos, abas **inexistentes** na destino são criadas automaticamente.

---

## 🧪 Testes

```bash
pip install -r requirements-dev.txt
pytest                                   # roda os 46 testes
pytest --cov=core --cov=config --cov-report=term-missing   # com cobertura
```

Cobertura atual: **~89%** (núcleo + config). Os testes geram os workbooks em
memória (não dependem do `RD_MODELO.xlsx` real) e validam valores, fórmulas,
estilos, mescladas, validações, formatação condicional, dimensões, comentários,
hyperlinks, gráficos, *named ranges*, os dois modos de cópia, criação de abas e
a análise de dependências — incluindo um *round-trip* salvar→reabrir.

---

## ✅ Teste manual end-to-end

**Opção rápida (sem o RD_MODELO real)** — usando o fixture gerado:
1. `streamlit run app.py`
2. Passo 1 → *Subir outro modelo* → envie `tests/fixtures/sample_template.xlsx`.
3. Passo 2 → envie **uma cópia** do mesmo arquivo como destino.
4. Passo 3 → selecione `Dados` (e `Calc`).
5. Passo 4 → observe os avisos de dependência (`FATURAS`, `Análise Econômica`),
   marque *“Entendo o risco”*.
6. Execute, baixe e abra no Excel: confira que `=SUM(B2:B3)`, a validação
   Sim/Não, a célula mesclada `A6:C6` e o gráfico foram preservados.

**Opção completa (com o RD_MODELO real):**
1. Coloque `RD_MODELO.xlsx` em `assets/` e use-o como **modelo padrão**.
2. Como **destino**, suba uma cópia do `RD_MODELO.xlsx` com alguns dados
   alterados.
3. Mantenha as **7 abas padrão** pré-selecionadas e execute a cópia.
4. Baixe e abra no Excel. Verifique especialmente:
   - fórmulas das abas `Análise Econômica` e `Tabela de Transferência`;
   - validações (listas) em `UC_DADOS`;
   - gráficos em `UC_DADOS` e `CONSUMO_EQP`;
   - os *named ranges* (`ListaSimNao`, `Modalidade`, `tipoCELPE`, `celpeCEE`).

---

## 🛠️ Troubleshooting

| Sintoma | Causa provável / solução |
|---|---|
| *“Modelo padrão não encontrado”* | Falta o `assets/RD_MODELO.xlsx`. Adicione-o ou use *Subir outro modelo*. |
| *“Arquivo Excel inválido”* | O arquivo está corrompido ou não é `.xlsx`/`.xlsm`. |
| *“Sem permissão para ler o arquivo”* | O arquivo está aberto no Excel. Feche-o e tente de novo. |
| Validação/lista “não funciona” na destino | A validação usa um *named range* que não existe na destino. A app copia os *named ranges* ausentes automaticamente — confirme que a aba que define o range também foi copiada. |
| Fórmulas aparecem como `0` ou vazias ao abrir | Normal: o Excel **recalcula ao abrir**. Salve no Excel uma vez. |
| Arquivo muito grande / memória | Feche outros apps; a app já trabalha em memória sem cópias extras em disco. |
| `streamlit: command not found` | Use `py -m streamlit run app.py` (Windows) ou ative o venv. |

Erros inesperados são exibidos na própria UI com um expander **“Detalhes
técnicos”** contendo o traceback.

---

## ⚠️ Limitações conhecidas

- **Recálculo de fórmulas**: a app não recalcula valores — o Excel o faz ao
  abrir o arquivo (comportamento intencional).
- **Referências 3D** em fórmulas (ex.: `=SUM('Aba1:Aba3'!A1)`) são tratadas como
  uma única referência na análise de dependências; podem gerar um aviso de
  dependência “quebrada” mesmo quando as abas existem. Não afeta a cópia.
- **Configuração de página** (margens, área de impressão) é copiada em modo
  *best-effort*: se algum atributo falhar, vira um aviso e a cópia continua.

---

## 🔒 Garantias

- Os arquivos de origem (modelo e destino) **não são alterados** — toda a
  manipulação ocorre sobre cópias em memória.
- O modelo é aberto **sem** `data_only`, preservando as fórmulas.
- Arquivos `.xlsm` são abertos e salvos com `keep_vba=True` (macros preservadas).
