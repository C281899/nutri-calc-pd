# CHANGELOG — NutriCalc P&D

## v2.2.0 — 2026-11-11

Auditoria completa com correção de erros de cálculo regulatório (Valores Diários,
arredondamento e quantidades não significativas da IN 75/2020), de leitura de
fichas, de backup e de desempenho. Lista completa em `AUDITORIA.md`.
Inclui suíte de testes (`python -m pytest tests -q`).

---

## v2.1.0 — 2026-05-22

### Resumo
Implementação das Etapas 1–4: suporte a entradas em mL e %, precisão de 4 casas decimais, indicador de soma percentual em tempo real, e tabela técnica P&D aprimorada no Excel.

---

### Arquivos Modificados

#### `modules/calculator.py`
- **Adicionado** `smart_round(valor, contexto)` — formata números para exibição sem zeros desnecessários (até 4 decimais) ou retorna `float` arredondado a 4 casas para cálculo
- **Adicionado** `convert_to_grams(valor, unidade, densidade, total_receita_g)` — converte mL (via densidade g/mL) e % (Modo A: base peso fixo; Modo B: base 100g) para gramas

#### `modules/database.py`
- **Adicionado** `_SQL_DENSIDADES` — DDL da tabela `densidades_ingredientes`
- **Adicionado** `_DENSIDADES_INICIAIS` — 18 densidades padrão (leite, óleos, mel, álcool, etc.)
- **Adicionado** `migrar_banco()` — migração idempotente que adiciona colunas novas em `receita_ingredientes` e cria/semeia `densidades_ingredientes`
- **Adicionado** `buscar_densidade(nome)` — fuzzy match (score ≥ 70) contra tabela de densidades
- **Adicionado** `salvar_densidade(nome, densidade, fonte)` — insert-or-update na tabela de densidades
- **Atualizado** `salvar_receita()` — persiste `unidade_original`, `quantidade_original`, `densidade_utilizada` em `receita_ingredientes`

#### `data/taco_setup.py`
- **Atualizado** `SQL_APP_DB` — adicionadas colunas `unidade_original TEXT DEFAULT 'g'`, `quantidade_original REAL DEFAULT NULL`, `densidade_utilizada REAL DEFAULT 1.0` à tabela `receita_ingredientes`
- **Adicionado** `CREATE TABLE IF NOT EXISTS densidades_ingredientes` ao `SQL_APP_DB`
- **Adicionado** `_DENSIDADES_INICIAIS` e seed inicial em `criar_app_db()`

#### `modules/excel_generator.py`
- **Atualizado** `_construir_aba_pd()`:
  - Colunas fixas expandidas de 4 para 6: adicionadas "Qtd. Original" e "Unidade" (cols D e E)
  - Precisão numérica aumentada de 3 para 4 casas decimais (`number_format = '0.0000'`)
  - Índices de colunas de nutrientes corrigidos (agora iniciam na coluna 7, fórmula `_NUT_START + j*2`)
  - Linha de totais e linha por 100g atualizadas com o novo layout de 6 colunas fixas

#### `app.py`
- **Adicionado** chamada a `db.migrar_banco()` na inicialização (migração automática de bancos existentes)
- **Adicionado** novos defaults de session state: `form_unit`, `form_modo_pct`, `form_total_receita_g`, `form_densidade_manual`
- **Modo Manual — seletor de unidade (g / mL / %):**
  - Layout alterado para 4 colunas: busca | quantidade | unidade | botão
  - `st.number_input` de quantidade: `step=0.0001, format="%.4f", min_value=0.0`
  - Ao selecionar **mL**: mostra densidade auto-detectada + campo para densidade manual
  - Ao selecionar **%**: mostra seletor Modo A/B, campo de peso total (Modo A) e indicador de soma percentual em tempo real
  - `calc.convert_to_grams()` converte antes de armazenar em `session_state`
  - Ingrediente salvo com campos extras: `unidade_original`, `quantidade_original`, `densidade_utilizada`
- **Tabela de ingredientes:**
  - Nova coluna "Qtd. original" (display only — exibe valor e unidade originais)
  - Coluna "Qtd. (g)": `step=0.0001, format="%.4f", min_value=0.0`
  - Indicador de soma % exibido quando há ingredientes em modo %
  - Peso total exibido com `calc.smart_round()`
- **Upload de ficha:** `min_value=0.1 → 0.0`, `step=0.0001`, `format="%.4f"`
- **Salvar Receita:** passa `unidade_original`, `quantidade_original`, `densidade_utilizada` ao banco

---

### Restrições Mantidas
- **Nenhuma alteração** nas regras de arredondamento da RDC 429/2020 (`aplicar_arredondamentos()` intocada)
- **Nenhuma alteração** nas queries SQL existentes — apenas adições com `DEFAULT`
- **Retrocompatibilidade total** com receitas já salvas (`unidade_original` assume `'g'` por padrão)

---

### Como Usar

#### Entrada em mL
1. Busque o ingrediente normalmente
2. Selecione "mL" no seletor de unidade
3. A densidade é detectada automaticamente se o ingrediente constar na tabela; ajuste se necessário
4. A quantidade é convertida para gramas antes do cálculo

#### Entrada em %
1. Selecione "%" no seletor de unidade
2. Escolha Modo A (peso total fixo) ou Modo B (base 100g)
3. O indicador de soma exibe o total de % acumulado em tempo real
4. A soma deve chegar a 100% para uma formulação completa

#### Precisão de 4 Decimais
- Todos os campos de quantidade aceitam até 4 casas decimais
- O Excel P&D exporta todos os valores nutricionais com 4 casas decimais

---

### Testes Realizados
- Banco existente migrado sem perda de dados (ALTER TABLE idempotente)
- `convert_to_grams()`: testado para g (identidade), mL×densidade, % Modo A, % Modo B
- `smart_round()`: testado com 0, inteiros, decimais com/sem zeros finais
- Excel gerado com 6 colunas fixas + colunas de nutrientes corretamente alinhadas
