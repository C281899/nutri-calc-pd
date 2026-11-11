# Auditoria do NutriCalc P&D — novembro/2026

Auditoria completa do código (app.py, módulos, scripts de dados e configuração de
deploy), com correção dos problemas encontrados. Base: commit `066ae86` (rebranding).

Como conferir: `python -m pytest tests -q` (49 testes) e `streamlit run app.py`.

## 1. Erros que alteravam o rótulo

| # | Problema encontrado | Correção |
|---|---|---|
| 1 | **Valores Diários desatualizados** — proteínas 75 g, gorduras saturadas 22 g, sódio 2.300 mg; gorduras trans sem %VD. | IN 75/2020, Anexo II: proteínas 50 g, saturadas 20 g, sódio 2.000 mg, trans 2 g. O %VD de proteína saía 33% menor que o correto; o de sódio, 13% menor. |
| 2 | **Arredondamento fora da norma** — tudo com 1 casa decimal e `round()` do Python (0,25 → 0,2). | Anexo III: ≥ 10 inteiro; de 1 a 10 uma casa; < 1 uma casa (g) ou duas (mg); sempre "meio para cima". |
| 3 | **Açúcares totais inventados** — o rótulo mostrava 40% dos carboidratos como se fosse dado. | Campo novo "Açúcares Totais na Receita"; sem ele a linha sai "Não informado". Nunca estimado. |
| 4 | **Quantidades não significativas** — só gorduras trans e sódio (com limite errado, 1 mg). | Anexo IV completo (energia ≤ 4 kcal, sódio ≤ 5 mg, etc.), valendo só quando o valor é insignificante em 100 g **e** na porção, e com as condições do anexo (gordura total não zera se houver saturada/trans; carboidrato não zera se houver açúcar). |
| 5 | **Validação TACO × TBCA trocava valores corretos** — aceitava alimento "parecido" com 60% de semelhança (ex.: leite em pó × leite fluido) e substituía ou tirava média dos valores da TACO. | A TBCA só preenche valor **ausente**; exige 88% de semelhança no nome e energia/umidade compatíveis; diferenças são apenas relatadas. |
| 6 | **Receita reaberta com outro alimento** — os ids mudam quando o banco é recriado (reinício do servidor, restauração de backup); a receita abria com a composição do alimento que herdou o id. | O id só vale se o nome conferir; senão a busca é pelo nome. Ingrediente não encontrado é avisado em vermelho. |
| 7 | **Quantidades embaralhadas ao remover um ingrediente** — os campos eram identificados pela posição na lista. | Cada linha tem identificador próprio. |
| 8 | **Resultado antigo na tela, no Excel e na receita salva** depois de mudar porção, nome ou ingredientes. | O resultado é refeito sempre que qualquer dado de entrada muda. |
| 9 | **Ingrediente sumia na importação de ficha** — "1.250,5 g" não era reconhecido e a linha era descartada sem aviso; "1.250" virava 1,25. | Leitura de número em formato brasileiro em todos os caminhos (PDF, Excel, CSV, TXT). Itens não importados são listados. |
| 10 | **Ficha de fornecedor lida errado** — "Gorduras saturadas" entrava como gorduras totais; sódio em gramas não era convertido; "Vitamina B12 2,4 mcg" virava 12; vírgula decimal quebrava colunas. | Corrigido; uma tabela nutricional também não é mais confundida com receita. |
| 11 | **Energia em kJ subestimada** quando a ficha do ingrediente só trazia kcal. | A unidade ausente é calculada (1 kcal = 4,184 kJ). |
| 12 | **Linha de sódio escondida** na tabela de resultado (atrás de barra de rolagem). | Tabela mostra todas as linhas. |
| 13 | Rodapé do Excel com o texto da norma antiga; porção "30 g (30 g)"; porção 37,5 g exibida como 38 g. | Texto da IN 75/2020; porção e medida caseira corretas. |

## 2. Perdas operacionais

- **Backup**: importar duas vezes duplicava tudo; as receitas perdiam o vínculo com
  ingredientes de fornecedor; gordura saturada/trans e açúcares não eram salvos
  com a receita (nem no backup). Corrigido — os campos são salvos, restaurados e
  o arquivo é gerado na hora do clique (antes saía desatualizado).
- **Dependências faltando no deploy**: `pymupdf` e `pytesseract` não estavam no
  `requirements.txt`, então o OCR de PDFs digitalizados nunca funcionou na nuvem
  (apesar de o tesseract ser instalado); sem `xlrd`, arquivos `.xls` davam erro.
- **Água não existia na base** — impossível formular bebidas, caldas e massas com
  o valor por 100 g correto. Incluída "Água potável" (bancos já criados são
  atualizados sozinhos).
- **Avisos que nunca apareciam**: mensagens de sucesso seguidas de recarga da
  página eram apagadas antes de serem vistas.
- **Formulário de fornecedor** não carregava os dados ao editar ou ao extrair uma
  ficha (mantinha o que estava digitado).
- **Densidade**: "suco de uva" herdava em silêncio a densidade de outro produto;
  o campo mantinha a densidade do ingrediente anterior. Agora mostra de qual
  item da tabela veio o valor.
- **Inicialização**: download da TACO sem tempo limite (podia travar a abertura);
  banco deixado pela metade por uma inicialização interrompida era aceito; dois
  acessos simultâneos criavam o banco duas vezes; consulta de 8 s ao site da
  TBCA cujo resultado era descartado.
- Conferências novas: saturada + trans maior que gorduras totais, açúcares
  maiores que carboidratos, porção maior que a receita, soma zero.
- Nome do arquivo Excel com `/ : ?` quebrava o download no Windows.
- Legendas (peso total, porção) quase ilegíveis no tema escuro.

## 3. Perdas computacionais

O Streamlit reexecuta o script inteiro a cada clique ou tecla. A cada execução o app:

| Antes | Agora |
|---|---|
| Rodava a migração do banco, com 3 `ALTER TABLE` que falhavam e eram ignorados | Uma vez por processo |
| Relia e normalizava todos os alimentos da TACO a cada busca | Índice em memória (busca ≈ 1 ms) |
| Relia a TBCA inteira para cada ingrediente de cada cálculo | Lida uma vez |
| Regerava o Excel completo | Gerado uma vez por resultado |
| Montava o backup JSON só para exibir o botão | Só no clique |
| Abria conexões sem fechar em caso de erro | Fechamento garantido |

## 4. O que continua pendente

1. **Armazenamento temporário** — no Streamlit Cloud os dados somem quando o
   servidor reinicia. O backup ameniza, mas a solução é um banco externo
   (ex.: Supabase/Postgres). É mudança de arquitetura, não foi feita aqui.
2. **Gordura saturada, trans e açúcares** continuam digitados por receita: a TACO
   não traz esses dados por alimento.
3. **Importação da planilha oficial da TACO** (caminho usado na nuvem, 1.082
   alimentos) não pôde ser executada neste ambiente (sem acesso ao site da
   UNICAMP); os testes rodaram com a base embutida (536 alimentos).
4. **Rotulagem frontal** (lupa de "alto em") e porções de referência por categoria
   (Anexo V) não são verificadas pelo app.
5. A validação por nutricionista habilitado antes do uso comercial continua obrigatória.

## 5. Atenção ao atualizar

- `requirements.txt` agora exige **Streamlit 1.53+**. Na instalação local rode
  `pip install -r requirements.txt --upgrade` antes de abrir o app.
- Os %VD de proteína, gordura saturada e sódio **mudam** em relação aos relatórios
  já emitidos. Excel gerado antes desta correção deve ser refeito.
- `CLAUDE.md` foi atualizado: trazia os Valores Diários antigos como regra.
