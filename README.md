# 🥗 NutriCalc P&D

Aplicação web para cálculo de tabelas nutricionais no padrão **ANVISA (RDC 429/2020)**.
Desenvolvida para uso interno em setores de Pesquisa & Desenvolvimento de alimentos e bebidas.

---

## ✨ Funcionalidades

- Cálculo automático de tabela nutricional com base nos dados da **TACO 4ª edição (UNICAMP)**
- Upload de fichas técnicas (PDF, Excel, CSV, TXT) com extração automática de ingredientes
- Cadastro de ingredientes industrializados de fornecedores
- Geração de planilha Excel profissional com duas abas:
  - **Tabela Nutricional — Rótulo** (layout fiel ao modelo ANVISA)
  - **Composição Técnica — P&D** (detalhamento interno completo)
- Histórico de receitas salvas com edição, duplicação e reprocessamento
- Interface 100% em português brasileiro, sem instalação local

---

## 🚀 Como Fazer o Deploy (Streamlit Community Cloud)

### Pré-requisitos
- Conta no [GitHub](https://github.com) (gratuita)
- Conta no [Streamlit Community Cloud](https://streamlit.io/cloud) (gratuita)

### Passo a Passo

**1. Fork do repositório**
```
GitHub → botão "Fork" → criar cópia no seu perfil
```

**2. (Opcional) Adicionar a planilha TACO completa**

Para obter todos os ~2.000 alimentos da TACO 4ª edição:
1. Acesse https://www.cfn.org.br/taco ou a UNICAMP diretamente
2. Baixe o arquivo Excel da TACO 4ª edição
3. Renomeie para `TACO_4ed.xlsx`
4. Coloque na pasta `data/` do repositório

> Sem este arquivo, a aplicação usa uma base embutida com os 50 alimentos mais comuns.

**3. Conectar ao Streamlit Cloud**
1. Acesse [share.streamlit.io](https://share.streamlit.io)
2. Clique em **"New app"**
3. Selecione seu fork do repositório
4. Branch: `main`
5. Main file path: `app.py`
6. Clique em **"Deploy!"**

**4. Aguardar o deploy**
O Streamlit Cloud instalará automaticamente todas as dependências do `requirements.txt`.
O banco de dados TACO será inicializado automaticamente na primeira execução.

**5. Acessar a aplicação**
Você receberá uma URL pública no formato:
```
https://seu-usuario-nutri-calc-pd.streamlit.app
```

---

## 🖥️ Rodar Localmente (Desenvolvimento)

### Requisitos
- Python 3.11 ou superior
- pip

### Instalação

```bash
# 1. Clonar o repositório
git clone https://github.com/seu-usuario/nutri-calc-pd.git
cd nutri-calc-pd

# 2. Criar ambiente virtual (recomendado)
python -m venv .venv
source .venv/bin/activate        # Linux/Mac
# ou
.venv\Scripts\activate           # Windows

# 3. Instalar dependências
pip install -r requirements.txt --upgrade

# 4. Inicializar o banco de dados (apenas na primeira vez)
python data/taco_setup.py

# 5. Rodar a aplicação
streamlit run app.py
```

A aplicação abrirá automaticamente em `http://localhost:8501`

---

## 📁 Estrutura do Projeto

```
nutri_calc_pd/
├── app.py                        # Interface principal (Streamlit)
├── requirements.txt              # Dependências Python fixadas
├── CLAUDE.md                     # Guia mestre para Claude Code
├── README.md                     # Este arquivo
│
├── modules/
│   ├── __init__.py
│   ├── database.py               # Operações SQLite (TACO + receitas + fornecedores)
│   ├── calculator.py             # Motor de cálculo nutricional (RDC 429/2020)
│   ├── excel_generator.py        # Geração do relatório Excel (openpyxl)
│   └── parser.py                 # Leitura de PDF, Excel, CSV, TXT
│
├── data/
│   ├── taco_setup.py             # Script de inicialização dos bancos
│   └── TACO_4ed.xlsx             # (opcional) Planilha TACO completa
│
└── .streamlit/
    └── config.toml               # Tema visual personalizado
```

**Bancos de dados (gerados automaticamente):**
```
taco.db         — Tabela TACO 4ª edição (alimentos + composição)
nutri_calc.db   — Receitas salvas + ingredientes de fornecedor
```

---

## 📋 Como Usar

### Calcular uma Receita

1. Acesse a aba **🥗 Nova Receita**
2. Preencha o nome do produto e o tamanho da porção
3. Adicione os ingredientes:
   - **Modo Manual**: busque pelo nome e informe a quantidade em gramas
   - **Modo Upload**: faça upload de uma ficha técnica (PDF, Excel, CSV ou TXT)
4. Clique em **🧮 Calcular Tabela Nutricional**
5. Baixe o Excel gerado ou salve a receita

### Cadastrar Ingrediente de Fornecedor

1. Acesse a aba **🏭 Ingredientes de Fornecedor**
2. Preencha o nome comercial e os valores nutricionais por 100g
   (com base na ficha técnica do seu fornecedor)
3. Clique em **💾 Salvar Ingrediente**
4. O ingrediente ficará disponível na busca da Nova Receita

### Reutilizar Receitas Salvas

1. Acesse a aba **📋 Receitas Salvas**
2. Clique em **📂 Abrir para Edição** para recarregar uma receita
3. Use **📋 Duplicar** para criar variações do mesmo produto
4. Edite e recalcule quantas vezes precisar

---

## ⚠️ Limitações Conhecidas

| Limitação | Detalhes |
|---|---|
| **Gordura Saturada e Trans** | Não diferenciadas na TACO. Devem ser informadas manualmente nos campos extras |
| **Açúcares Adicionados** | A TACO não diferencia açúcares adicionados de naturais. Informar manualmente |
| **Processamento Térmico** | Perdas nutricionais por cozimento/pasteurização não são calculadas |
| **Variações de Safra** | A TACO fornece valores médios; variações por origem/safra não são contempladas |
| **Validação Obrigatória** | ⚠️ **Obrigatório**: validação por nutricionista habilitado (CRN) antes de uso em rótulo comercial |
| **Ingredientes Industrializados** | Aditivos, concentrados e extratos não constam na TACO. Cadastre na aba de Fornecedores |

---

## 🏛️ Base Legal

- **RDC 429/2020 (ANVISA)**: Regulamento Técnico sobre Informação Nutricional em Rótulos de Alimentos
- **TACO 4ª edição**: Tabela Brasileira de Composição de Alimentos — UNICAMP, 2011
- **Valores Diários de Referência**: Adultos (2.000 kcal)

---

## 🛠️ Desenvolvido com Claude Code

Este projeto foi gerado e estruturado com o auxílio do [Claude Code](https://claude.ai/code)
da Anthropic, seguindo boas práticas de engenharia de software para aplicações Python/Streamlit.

---

## 📄 Licença

Uso interno — P&D. Consulte o responsável técnico antes de redistribuir.
