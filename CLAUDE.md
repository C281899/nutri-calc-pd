# NutriCalc P&D — Guia Mestre para Claude Code

## Visão Geral do Projeto
Aplicação web completa chamada **NutriCalc P&D** para uso interno em setor de
Pesquisa & Desenvolvimento de alimentos e bebidas. Gera tabelas nutricionais no
padrão ANVISA (RDC 429/2020) a partir de receitas digitadas ou importadas.

## Stack Técnica
- **Frontend/Backend**: Streamlit (Python 3.11+)
- **Banco de dados**: SQLite local (`taco.db` + `nutri_calc.db`)
- **Parsing de arquivos**: pdfplumber, openpyxl, pandas
- **Geração de Excel**: openpyxl com formatação profissional
- **Matching**: RapidFuzz (fuzzy matching, threshold 80%)
- **Deploy**: Streamlit Community Cloud (gratuito)

## Estrutura de Arquivos Obrigatória
```
nutri_calc_pd/
├── app.py                    # Entry point Streamlit — orquestra abas e sessão
├── modules/
│   ├── __init__.py
│   ├── database.py           # Toda interação com SQLite
│   ├── calculator.py         # Motor de cálculo nutricional (RDC 429/2020)
│   ├── excel_generator.py    # Geração do relatório Excel (openpyxl)
│   └── parser.py             # Leitura de PDF, Excel, CSV, TXT
├── data/
│   └── taco_setup.py         # Script único de importação da TACO para SQLite
├── .streamlit/
│   └── config.toml           # Tema visual personalizado
├── requirements.txt          # Dependências fixadas
├── README.md                 # Instruções de deploy
└── CLAUDE.md                 # Este arquivo
```

---

## ETAPA 1 — Banco de Dados (database.py + taco_setup.py)

### taco_setup.py
Script executado **uma única vez** na primeira inicialização. Deve:
1. Baixar a planilha TACO 4ª edição (Excel) do servidor da UNICAMP ou usar
   arquivo local `data/TACO_4ed.xlsx` se já presente.
2. Criar `taco.db` com a tabela `alimentos_taco`:

```sql
CREATE TABLE alimentos_taco (
    id INTEGER PRIMARY KEY,
    nome_alimento TEXT NOT NULL,
    categoria TEXT,
    umidade REAL, energia_kcal REAL, energia_kj REAL,
    proteina REAL, lipideos REAL, colesterol REAL,
    carboidrato REAL, fibra_alimentar REAL, cinzas REAL,
    calcio REAL, magnesio REAL, manganes REAL, fosforo REAL,
    ferro REAL, sodio REAL, potassio REAL, cobre REAL, zinco REAL,
    retinol REAL, re REAL, rae REAL, tiamina REAL, riboflavina REAL,
    piridoxina REAL, niacina REAL, vitamina_c REAL, vitamina_d REAL,
    vitamina_e REAL, vitamina_b12 REAL
);
```

3. Criar tabela `ingredientes_fornecedor`:
```sql
CREATE TABLE ingredientes_fornecedor (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nome_comercial TEXT NOT NULL,
    nome_generico TEXT,
    fabricante TEXT,
    -- todos os nutrientes acima (mesma estrutura) --
    data_cadastro TEXT DEFAULT (datetime('now'))
);
```

4. Criar banco `nutri_calc.db` com tabelas de receitas:
```sql
CREATE TABLE receitas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nome_produto TEXT NOT NULL,
    porcao_gramas REAL NOT NULL,
    num_porcoes INTEGER,
    data_criacao TEXT DEFAULT (datetime('now')),
    data_atualizacao TEXT
);

CREATE TABLE receita_ingredientes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    receita_id INTEGER REFERENCES receitas(id) ON DELETE CASCADE,
    nome_ingrediente TEXT NOT NULL,
    fonte TEXT NOT NULL,  -- 'TACO' ou 'FORNECEDOR'
    fonte_id INTEGER,     -- id na tabela de origem
    quantidade_gramas REAL NOT NULL,
    eh_subrecita INTEGER DEFAULT 0,
    subrecita_id INTEGER REFERENCES receitas(id)
);
```

### database.py — Funções obrigatórias:
```python
def buscar_ingrediente(termo: str, limite=10) -> list[dict]
    # Fuzzy match contra taco + fornecedor, retorna lista de candidatos
    # com campos: id, nome, fonte, score

def get_composicao_por_100g(fonte: str, id: int) -> dict
    # Retorna dict com todos os nutrientes por 100g

def salvar_receita(nome, porcao, ingredientes, num_porcoes=None) -> int
    # Persiste receita, retorna id

def listar_receitas() -> list[dict]
    # Lista todas as receitas salvas

def get_receita_completa(receita_id: int) -> dict
    # Retorna receita com todos os ingredientes e composições

def salvar_ingrediente_fornecedor(dados: dict) -> int
    # Cadastra novo ingrediente industrializado

def listar_ingredientes_fornecedor() -> list[dict]

def deletar_ingrediente_fornecedor(id: int)

def deletar_receita(id: int)
```

---

## ETAPA 2 — Calculator (calculator.py)

### Constantes VD (IN 75/2020, Anexo II):
```python
VALORES_DIARIOS = {
    "energia_kcal": 2000,
    "energia_kj": 8400,
    "carboidrato": 300,
    "acucares_adicionados": 50,
    "proteina": 50,
    "lipideos": 65,
    "gordura_saturada": 20,
    "gordura_trans": 2,
    "fibra_alimentar": 25,
    "sodio": 2000,
}
```

### Regras de Arredondamento (IN 75/2020, Anexo III):
```python
# Valor >= 10 ............ número inteiro
# 1 <= valor < 10 ........ uma casa decimal
# valor < 1 (g, kcal) .... uma casa decimal
# valor < 1 (mg, mcg) .... duas casas decimais
# Sempre "meio para cima" (0,25 → 0,3) — NÃO usar round() do Python.
# %VD: inteiro.
# Quantidades não significativas (Anexo IV) viram zero só quando o valor é
# insignificante em 100 g E na porção: energia <= 4 kcal; carboidratos,
# açúcares, proteínas, gorduras totais e fibras <= 0,5 g; saturadas e trans
# <= 0,1 g; sódio <= 5 mg.
# Açúcares totais NUNCA são estimados: sem dado informado → "Não informado".
```

### Funções obrigatórias:
```python
def calcular_composicao_receita(ingredientes: list[dict]) -> dict
    """
    ingredientes = [
        {"composicao_100g": {...}, "quantidade_gramas": float},
        ...
    ]
    Retorna composição TOTAL da receita (soma ponderada).
    """

def calcular_por_porcao(composicao_total: dict, porcao_gramas: float) -> dict
    """Escala composição total para a porção informada."""

def calcular_por_100g_produto(composicao_total: dict, peso_total_receita: float) -> dict
    """Normaliza composição total para base de 100g."""

def calcular_vd(por_porcao: dict) -> dict
    """Retorna %VD para cada nutriente obrigatório."""

def aplicar_arredondamentos(valores: dict) -> dict
    """Aplica regras de arredondamento da ANVISA."""

def resolver_subrecita(subrecita_id: int, quantidade_gramas: float) -> dict
    """Resolve recursivamente sub-receitas, retorna composição por 100g."""
```

---

## ETAPA 3 — Parser (parser.py)

### Formatos suportados:
- **PDF**: pdfplumber — extrair tabelas e texto, regex para padrões "ingrediente: Xg"
- **Excel/CSV**: pandas — detectar colunas de ingrediente e quantidade
- **TXT**: regex e heurísticas

### Função principal:
```python
def parse_ficha_tecnica(arquivo_bytes: bytes, extensao: str) -> list[dict]:
    """
    Retorna lista de: [{"nome": str, "quantidade_gramas": float}, ...]
    Lança ValueError com mensagem amigável se não conseguir extrair.
    """
```

### Estratégia de extração:
1. Tentar extrair tabelas estruturadas primeiro
2. Se falhar, procurar padrões: `"ingrediente X: Yg"`, `"X (Yg)"`, linhas com número seguido de 'g'
3. Retornar dados brutos para confirmação pelo usuário (nunca silencioso)

---

## ETAPA 4 — Excel Generator (excel_generator.py)

### Aba 1 — "Tabela Nutricional — Rótulo" (layout ANVISA):
```
INFORMAÇÃO NUTRICIONAL
Porção de Xg (X medida caseira)
% Valores Diários com base em uma dieta de 2.000 kcal ou 8.400 kJ.

| Nutriente              | Por 100g | Por Porção (Xg) | %VD* |
|------------------------|----------|------------------|------|
| Valor Energético       | X kcal   | X kcal           | X%   |
|                        | X kJ     | X kJ             |      |
| Carboidratos           | Xg       | Xg               | X%   |
| Açúcares Totais        | Xg       | Xg               | **   |
| Açúcares Adicionados   | Xg       | Xg               | X%   |
| Proteínas              | Xg       | Xg               | X%   |
| Gorduras Totais        | Xg       | Xg               | X%   |
| Gorduras Saturadas     | Xg       | Xg               | X%   |
| Gorduras Trans         | Xg       | Xg               | **   |
| Fibra Alimentar        | Xg       | Xg               | X%   |
| Sódio                  | Xmg      | Xmg              | X%   |

* % Valores Diários de referência com base em uma dieta de 2.000 kcal...
** Não estabelecido.
```

### Formatação openpyxl:
- Fonte: Arial 10pt (dados), Arial 11pt bold (cabeçalhos)
- Cabeçalho "INFORMAÇÃO NUTRICIONAL": merge de células, negrito, fundo #1F4E79, fonte branca
- Zebra striping: linhas alternadas #F2F2F2 / branco
- Bordas: thin em todas as células da tabela
- Larguras fixas: coluna A=35, B=15, C=18, D=10

### Aba 2 — "Composição Técnica — P&D":
- Todos os ingredientes com suas quantidades e contribuições individuais
- Todos os micronutrientes da TACO (não só os obrigatórios ANVISA)
- Coluna de % de contribuição de cada ingrediente no total
- Linha de TOTAIS ao final, em negrito

```python
def gerar_excel(
    nome_produto: str,
    porcao_gramas: float,
    num_porcoes: int | None,
    ingredientes_detalhes: list[dict],
    composicao_por_100g: dict,
    por_porcao: dict,
    vd: dict,
) -> bytes:
    """Retorna bytes do arquivo Excel para download."""
```

---

## ETAPA 5 — app.py (Interface Streamlit)

### Inicialização:
```python
# Verificar se taco.db existe, se não → rodar taco_setup.py automaticamente
# Inicializar st.session_state com:
#   - ingredientes_atuais: []
#   - modo_entrada: "manual"
#   - receita_em_edicao: None
```

### Estrutura de abas:
```python
tab1, tab2, tab3 = st.tabs(["🥗 Nova Receita", "🏭 Ingredientes de Fornecedor", "📋 Receitas Salvas"])
```

### Aba 1 — Nova Receita:
```
st.text_input("Nome do Produto *")
st.number_input("Tamanho da Porção (g) *", min_value=1.0)
st.number_input("Número de Porções por Embalagem", min_value=1)

# Botão alternância modo
modo = st.radio("Modo de entrada", ["✏️ Manual", "📎 Upload de Ficha"])

# MODO MANUAL:
#   - st.text_input com busca dinâmica (st.selectbox + buscar_ingrediente)
#   - st.number_input quantidade
#   - st.button "+ Adicionar")
#   - st.dataframe(ingredientes_atuais, hide_index=True) com coluna de remover

# MODO UPLOAD:
#   - st.file_uploader(accept ["pdf","xlsx","xls","csv","txt"])
#   - Ao fazer upload: chamar parser.parse_ficha_tecnica()
#   - st.data_editor() para edição da prévia
#   - Para cada linha: st.selectbox com matches do fuzzy

# Botão "🧮 Calcular Composição Nutricional"
# Validações antes de calcular:
#   - Nome preenchido?
#   - Porção > 0?
#   - Pelo menos 1 ingrediente?
# Ao calcular:
#   - Chamar calculator.calcular_*()
#   - Exibir tabela de resultados na tela
#   - st.download_button com gerar_excel()
#   - Botão "💾 Salvar Receita"
```

### Aba 2 — Ingredientes de Fornecedor:
```
# Formulário de cadastro com todos os campos nutricionais
# Upload opcional de ficha do fornecedor
# Tabela de ingredientes cadastrados (st.dataframe com botões editar/excluir)
```

### Aba 3 — Receitas Salvas:
```
# st.dataframe com lista de receitas
# Para cada receita: botões Abrir, Duplicar, Excluir
# Abrir → carrega na Aba 1 para reedição
```

---

## ETAPA 6 — Arquivos de Configuração

### .streamlit/config.toml:
```toml
[theme]
primaryColor = "#1F4E79"
backgroundColor = "#FFFFFF"
secondaryBackgroundColor = "#F0F4F8"
textColor = "#1A1A2E"
font = "sans serif"

[server]
maxUploadSize = 50

[browser]
gatherUsageStats = false
```

### requirements.txt:
Ver o arquivo `requirements.txt` (Streamlit 1.53+, pymupdf e pytesseract para
OCR, xlrd para .xls).

### Testes:
```bash
python -m pytest tests -q
```
Rodar sempre antes de publicar. `AUDITORIA.md` lista o que já foi corrigido e
por quê — não reintroduzir.

---

## Mensagens de Erro Padrão (Interface)
```python
# Usar st.error() para erros, st.warning() para avisos, st.success() para confirmações
# NUNCA expor stacktrace ao usuário
# Padrão: "❌ [O que aconteceu]. [O que o usuário deve fazer]."

ERROS = {
    "sem_ingredientes": "❌ Adicione pelo menos um ingrediente antes de calcular.",
    "sem_porcao": "❌ Informe o tamanho da porção em gramas.",
    "sem_nome": "❌ Informe o nome do produto.",
    "ingrediente_nao_encontrado": "⚠️ Ingrediente '{nome}' não encontrado na TACO nem em fornecedores. Cadastre-o na aba 'Ingredientes de Fornecedor'.",
    "parse_falhou": "❌ Não foi possível extrair ingredientes do arquivo. Verifique se o arquivo está no formato correto ou use a entrada manual.",
    "db_erro": "❌ Erro ao acessar o banco de dados. Tente novamente.",
}
```

---

## LIMITAÇÕES CONHECIDAS
```python
# 1. Ingredientes não encontrados na TACO exigem cadastro manual de fornecedor
# 2. Variações de composição nutricional por safra/origem não cobertas pela TACO
# 3. OBRIGATÓRIO: validação por nutricionista habilitado antes de uso em rótulo comercial
# 4. Perdas nutricionais por processamento térmico não calculadas automaticamente
#    (TACO fornece valores de alimentos crus/padrão)
# 5. Açúcares Adicionados precisam ser informados manualmente (TACO não diferencia)
# 6. Medida caseira deve ser informada manualmente pelo usuário
```

---

## Ordem de Implementação Recomendada para Claude Code

1. `requirements.txt` → instalar dependências
2. `data/taco_setup.py` → criar e popular os bancos de dados
3. `modules/database.py` → funções de acesso aos dados
4. `modules/calculator.py` → motor de cálculo
5. `modules/parser.py` → leitura de arquivos
6. `modules/excel_generator.py` → geração do Excel
7. `app.py` → interface Streamlit completa
8. `.streamlit/config.toml` → tema visual
9. `README.md` → instruções de deploy
10. Testar end-to-end com uma receita simples

## Como Rodar Localmente
```bash
pip install -r requirements.txt
python data/taco_setup.py   # apenas na primeira vez
streamlit run app.py
```
