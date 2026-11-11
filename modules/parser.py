"""
parser.py — Leitura e extração de ingredientes de arquivos
Suporta: PDF, Excel (.xlsx/.xls), CSV, TXT
"""

import os
import re
import io
import pandas as pd
from typing import Any

# Detecta qual backend PDF está disponível (usado para info na UI)
try:
    import fitz as _fitz_check  # noqa: F401
    PDF_BACKEND = "pymupdf"
except ImportError:
    PDF_BACKEND = "pdfplumber"


def parse_ficha_tecnica(arquivo_bytes: bytes, extensao: str) -> list[dict]:
    """
    Ponto de entrada principal. Extrai ingredientes e quantidades de um arquivo.

    Retorna lista de dicts. Dois formatos possíveis:
      - Receita multi-ingredientes: [{"nome": str, "quantidade_gramas": float}, ...]
      - Ficha de ingrediente único:  [{"nome": str, "quantidade_gramas": 100,
                                       "composicao_direta": {nutrientes...}}]

    Raises:
        ValueError: com mensagem amigável se não conseguir extrair dados
    """
    ext = extensao.lower().strip(".")

    resultado: list[dict] = []
    try:
        if ext == "pdf":
            resultado = _parse_pdf(arquivo_bytes)
        elif ext in ("xlsx", "xls"):
            resultado = _parse_excel(arquivo_bytes, ext)
        elif ext == "csv":
            resultado = _parse_csv(arquivo_bytes)
        elif ext == "txt":
            resultado = _parse_txt(arquivo_bytes)
        else:
            raise ValueError(f"Formato '{ext}' não suportado. Use PDF, Excel, CSV ou TXT.")
    except ValueError:
        raise
    except Exception:
        resultado = []

    # Uma tabela nutricional (linhas "Proteínas 20 g", "Sódio 20 mg"...) não é
    # uma receita: se a maioria dos "ingredientes" lidos são nomes de nutrientes,
    # o arquivo é a ficha técnica de UM ingrediente.
    if resultado:
        n_nutrientes = sum(
            1 for r in resultado if _campo_do_label(str(r["nome"]).lower())
        )
        if not (n_nutrientes >= 3 and n_nutrientes * 2 >= len(resultado)):
            return resultado

    # ── Ficha de ingrediente único (composição por 100g) ──
    try:
        composicao = parse_composicao_nutricional(arquivo_bytes, ext)
        # Remove nome_comercial do dict de nutrientes, usa como nome
        nome = composicao.pop("nome_comercial", None)
        if not nome:
            # Extrair nome do arquivo não é possível aqui; deixar vazio para o usuário preencher
            nome = ""
        # Só aceitar se extraiu pelo menos 3 nutrientes reais
        nut_count = sum(1 for v in composicao.values() if v is not None)
        if nut_count >= 3:
            return [{"nome": nome, "quantidade_gramas": 100.0, "composicao_direta": composicao}]
    except (ValueError, Exception):
        pass

    raise ValueError(
        "❌ Não foi possível extrair ingredientes do arquivo. "
        "Verifique se o formato é uma lista de ingredientes com quantidades "
        "ou uma tabela nutricional por 100g."
    )


# ─── Padrões de extração (do mais ao menos específico) ───────────────────────
# Número: aceita milhar e decimal juntos ("1.250,5"). Com o padrão anterior o
# número era cortado em "1.250", a linha deixava de casar e o ingrediente
# sumia da receita sem aviso.
PADROES_QUANTIDADE = [
    # "Farinha de trigo (200g)" ou "Farinha de trigo (200 g)"
    r"([A-Za-zÀ-ÿ][^(\n]{2,70}?)\s*\(\s*(\d+(?:[.,]\d+)*)\s*(g|gramas?|kg|ml|l)\s*\)",
    # "Farinha de trigo: 200g" — separador : ou – ou -
    r"([A-Za-zÀ-ÿ][^:\-–\n]{2,70}?)\s*[:–\-]+\s*(\d+(?:[.,]\d+)*)\s*(g|gramas?|kg|ml|l)\b",
    # "200g Farinha de trigo" — quantidade no início
    r"^(\d+(?:[.,]\d+)*)\s*(g|gramas?|kg)\b[\s:–\-]+([A-Za-zÀ-ÿ][^\n,;]{2,70}?)\s*$",
    # Pontos entre nome e quantidade: "Farinha ........ 200 g"
    r"([A-Za-zÀ-ÿ][^.\n]{2,70}?)\.{2,}\s*(\d+(?:[.,]\d+)*)\s*(g|gramas?|kg|ml)?\s*$",
    # Tabela tab/pipe: "Farinha de trigo \t 200 g"
    r"([A-Za-zÀ-ÿ][^\t|\n]{2,70}?)[\t|]\s*(\d+(?:[.,]\d+)*)\s*(g|gramas?|kg)?\s*$",
    # Espaço duplo entre nome e número: "Farinha de trigo   200"
    r"([A-Za-zÀ-ÿ][a-zA-ZÀ-ÿ,. ()]{3,70}?)\s{2,}(\d+(?:[.,]\d+)*)\s*(g|gramas?|kg|ml)?\s*$",
    # Vírgula como separador: "Farinha de trigo, 200g"
    r"([A-Za-zÀ-ÿ][^,\n\d]{2,70}?),\s*(\d+(?:[.,]\d+)*)\s*(g|gramas?|kg|ml|l)\b",
    # Nome seguido da quantidade no fim da linha: "Sal 2,5 g", "Ovos 3 unidades 150g"
    r"^([A-Za-zÀ-ÿ][^\n:;|\t]{1,70}?)\s+(\d+(?:[.,]\d+)*)\s*(g|gramas?|kg|ml|l)\s*$",
]


def _num_br(texto) -> float:
    """
    Converte texto numérico para float aceitando o formato brasileiro.
    "200,5" → 200.5 | "1.250,00" → 1250.0 | "2.5" → 2.5 | "1.250" → 1250.0
    Lança ValueError se não houver número.
    """
    t = str(texto).strip().replace(" ", "")
    m = re.search(r"-?\d[\d.,]*", t)
    if not m:
        raise ValueError(f"sem número em {texto!r}")
    t = m.group().rstrip(".,")
    if "," in t and "." in t:
        if t.rfind(",") > t.rfind("."):          # 1.250,00
            t = t.replace(".", "").replace(",", ".")
        else:                                     # 1,250.00
            t = t.replace(",", "")
    elif "," in t:
        t = t.replace(",", ".")
    elif re.fullmatch(r"-?\d{1,3}(\.\d{3})+", t):  # 1.250 (milhar)
        t = t.replace(".", "")
    return float(t)


def _normalizar_quantidade(valor_str: str, unidade: str = "g") -> float:
    """Converte valor de quantidade para gramas (kg e litros ×1000; mL ≈ g)."""
    val = _num_br(valor_str)
    unidade = (unidade or "g").lower()
    if unidade in ("kg", "quilograma", "quilogramas", "l", "litro", "litros"):
        val *= 1000
    return val


def _limpar_nome(nome: str) -> str:
    """Remove artefatos comuns de extração de PDF/texto."""
    nome = re.sub(r'\s+', ' ', nome)
    nome = nome.strip(" \t\n\r-–•*:.,/\\")
    # Remover sufixos numéricos isolados (números de página)
    nome = re.sub(r'\s+\d{1,3}$', '', nome)
    return nome


def _extrair_por_regex(texto: str) -> list[dict]:
    """
    Tenta extrair ingredientes via múltiplos padrões regex.
    Aplica todos os padrões e desduplicata pelo par (nome, quantidade).
    """
    resultados = []
    vistos = set()
    trechos_usados: list[tuple[int, int]] = []

    for padrao in PADROES_QUANTIDADE:
        for match in re.finditer(padrao, texto, re.IGNORECASE | re.MULTILINE):
            ini, fim = match.span()
            if any(ini < f and fim > i for i, f in trechos_usados):
                continue  # trecho já interpretado por um padrão mais específico
            grupos = match.groups()
            try:
                if len(grupos) >= 3:
                    g0 = (grupos[0] or "").strip()
                    g1 = (grupos[1] or "").strip()
                    g2 = (grupos[2] or "g").strip()

                    # Detectar se é "qty unidade nome" (g0 começa com dígito)
                    if re.match(r'^\d', g0):
                        nome    = g2
                        qtd_str = g0
                        unidade = g1
                    else:
                        nome    = g0
                        qtd_str = g1
                        unidade = g2 or "g"
                elif len(grupos) == 2:
                    nome    = (grupos[0] or "").strip()
                    qtd_str = (grupos[1] or "").strip()
                    unidade = "g"
                else:
                    continue

                nome = _limpar_nome(nome)
                if len(nome) < 3 or not any(c.isalpha() for c in nome):
                    continue

                qtd = _normalizar_quantidade(qtd_str, unidade)
                if qtd <= 0 or qtd > 100_000:
                    continue

                chave = f"{nome.lower().strip()}:{qtd}"
                if chave not in vistos:
                    vistos.add(chave)
                    trechos_usados.append((ini, fim))
                    resultados.append({"nome": nome, "quantidade_gramas": qtd})

            except (ValueError, AttributeError, IndexError):
                continue

    return resultados


def _extrair_liberal(texto: str) -> list[dict]:
    """
    Fallback liberal: captura qualquer linha com texto + número plausível como grama.
    Usado quando os padrões estritos não encontram nada.
    """
    resultados = []
    vistos     = set()

    for linha in texto.splitlines():
        linha = linha.strip()
        if len(linha) < 4:
            continue

        nums = re.findall(r'([\d]+[.,]?\d*)', linha)
        if not nums:
            continue

        texto_linha = re.sub(r'\b[\d]+[.,]?\d*\s*(g|gramas?|kg|ml|l|mg)?\b', '', linha, flags=re.IGNORECASE)
        texto_linha = _limpar_nome(texto_linha)

        if len(texto_linha) < 3 or not any(c.isalpha() for c in texto_linha):
            continue

        for num_str in nums:
            try:
                qtd = float(num_str.replace(',', '.'))
                if 0.1 <= qtd <= 50000:
                    chave = f"{texto_linha.lower()}:{qtd}"
                    if chave not in vistos:
                        vistos.add(chave)
                        resultados.append({"nome": texto_linha, "quantidade_gramas": qtd})
                    break
            except ValueError:
                continue

    return resultados


def _parse_pdf(conteudo: bytes) -> list[dict]:
    """Extrai ingredientes de PDF usando pdfplumber com múltiplas estratégias."""
    try:
        import pdfplumber
    except ImportError:
        raise ValueError("pdfplumber não instalado. Execute: pip install pdfplumber")

    resultados_tabela = []
    texto_completo    = ""

    try:
        with pdfplumber.open(io.BytesIO(conteudo)) as pdf:
            for pagina in pdf.pages:
                # Estratégia 1: tabelas estruturadas
                tabelas = pagina.extract_tables()
                for tabela in tabelas:
                    resultados_tabela.extend(_processar_tabela(tabela))

                # Estratégia 2: texto da página
                texto = pagina.extract_text(x_tolerance=3, y_tolerance=3) or ""
                texto_completo += texto + "\n"

                # Estratégia 3: palavras com coordenadas (preserva espaçamento real)
                try:
                    words = pagina.extract_words(keep_blank_chars=False)
                    if words:
                        linha_por_y: dict[int, list] = {}
                        for w in words:
                            y_key = round(w["top"] / 5) * 5
                            linha_por_y.setdefault(y_key, []).append(w["text"])
                        for y in sorted(linha_por_y):
                            texto_completo += " ".join(linha_por_y[y]) + "\n"
                except Exception:
                    pass

    except Exception as e:
        raise ValueError(
            f"Erro ao ler PDF: {e}. "
            "Verifique se o arquivo não está corrompido ou protegido por senha."
        )

    # Prioridade: tabelas → regex → fallback liberal
    if resultados_tabela:
        return _deduplicar(resultados_tabela)

    resultados_regex = _extrair_por_regex(texto_completo)
    if resultados_regex:
        return _deduplicar(resultados_regex)

    resultados_liberal = _extrair_liberal(texto_completo)
    if resultados_liberal:
        return _deduplicar(resultados_liberal)

    raise ValueError(
        "Não foi possível extrair ingredientes do PDF. "
        "Possíveis causas: (1) PDF digitalizado como imagem, (2) proteção de senha, "
        "(3) formato não reconhecido. "
        "Dica: copie o texto do PDF e cole em um arquivo .txt, ou use a entrada manual."
    )


def _parse_excel(conteudo: bytes, ext: str) -> list[dict]:
    """Extrai ingredientes de planilha Excel."""
    try:
        engine = "openpyxl" if ext == "xlsx" else "xlrd"
        df = pd.read_excel(io.BytesIO(conteudo), engine=engine, header=None)
    except Exception as e:
        raise ValueError(f"Erro ao ler Excel: {e}")

    resultados = _detectar_colunas_excel(df)

    if not resultados:
        texto = df.to_csv(sep="\t", index=False)
        resultados = _extrair_por_regex(texto)

    if not resultados:
        raise ValueError(
            "Não foi possível identificar colunas de ingrediente e quantidade na planilha. "
            "Verifique se a planilha tem colunas com nome do ingrediente e quantidade em gramas."
        )

    return _deduplicar(resultados)


def _detectar_colunas_excel(df: pd.DataFrame) -> list[dict]:
    """Tenta identificar automaticamente colunas de ingrediente e quantidade."""
    PALAVRAS_INGREDIENTE = ["ingrediente", "matéria", "materia", "produto", "componente",
                            "item", "alimento", "descrição", "descricao", "nome"]
    PALAVRAS_QUANTIDADE  = ["quantidade", "qtd", "qty", "peso", "gramas", "g ", "massa",
                            "amount", "quant"]

    col_ing, col_qtd = None, None

    for i in range(min(5, len(df))):
        row = df.iloc[i].astype(str).str.lower()
        for j, cell in enumerate(row):
            if any(p in cell for p in PALAVRAS_INGREDIENTE):
                col_ing = j
            if any(p in cell for p in PALAVRAS_QUANTIDADE):
                col_qtd = j

        if col_ing is not None and col_qtd is not None:
            resultados = []
            for k in range(i + 1, len(df)):
                try:
                    nome = str(df.iloc[k, col_ing]).strip()
                    qtd  = _num_br(df.iloc[k, col_qtd])
                    if nome and nome.lower() not in ("nan", "") and qtd > 0:
                        resultados.append({"nome": nome, "quantidade_gramas": qtd})
                except (ValueError, IndexError):
                    continue
            if resultados:
                return resultados

    # Heurística posicional: coluna com texto + coluna com números
    if len(df.columns) >= 2:
        for col_texto in df.columns:
            for col_num in df.columns:
                if col_texto == col_num:
                    continue
                serie_texto = df[col_texto].astype(str)
                serie_num   = df[col_num].map(_num_ou_nan)

                if serie_texto.str.len().gt(2).sum() >= 2 and serie_num.notna().sum() >= 2:
                    resultados = []
                    for i in df.index:
                        try:
                            nome = str(df.at[i, col_texto]).strip()
                            qtd  = serie_num.at[i]
                            if nome and nome.lower() not in ("nan", "") and not pd.isna(qtd) and qtd > 0:
                                resultados.append({"nome": nome, "quantidade_gramas": float(qtd)})
                        except Exception:
                            continue
                    if len(resultados) >= 2:
                        return resultados

    return []


def _num_ou_nan(valor) -> float:
    """_num_br que devolve NaN para células vazias ou com texto (nomes)."""
    t = str(valor).strip()
    if not t or t.lower() == "nan" or re.search(r"[A-Za-zÀ-ÿ]{3,}", t):
        return float("nan")
    try:
        return _num_br(t)
    except ValueError:
        return float("nan")


def _parse_csv(conteudo: bytes) -> list[dict]:
    """Extrai ingredientes de CSV."""
    for encoding in ("utf-8", "utf-8-sig", "latin-1", "cp1252"):
        try:
            texto = conteudo.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        texto = conteudo.decode("utf-8", errors="replace")

    separador = ","
    for sep in (";", "\t", "|"):
        if texto.count(sep) > texto.count(","):
            separador = sep
            break

    try:
        # header=None + texto: o detector procura o cabeçalho nas primeiras linhas
        df = pd.read_csv(io.StringIO(texto), sep=separador, header=None,
                         dtype=str, keep_default_na=False, engine="python")
        resultados = _detectar_colunas_excel(df)
    except Exception:
        resultados = []

    if not resultados:
        resultados = _extrair_por_regex(texto)

    if not resultados:
        raise ValueError(
            "Não foi possível extrair ingredientes do CSV. "
            "Verifique se o arquivo tem colunas de ingrediente e quantidade separadas "
            "por vírgula ou ponto-e-vírgula."
        )

    return _deduplicar(resultados)


def _parse_txt(conteudo: bytes) -> list[dict]:
    """Extrai ingredientes de texto livre."""
    for encoding in ("utf-8", "utf-8-sig", "latin-1", "cp1252"):
        try:
            texto = conteudo.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        texto = conteudo.decode("utf-8", errors="replace")

    resultados = _extrair_por_regex(texto)

    if not resultados:
        resultados = _extrair_liberal(texto)

    if not resultados:
        raise ValueError(
            "Não foi possível identificar ingredientes e quantidades no arquivo de texto. "
            "Verifique se o arquivo segue o padrão 'Nome do Ingrediente: quantidade g' ou similar."
        )

    return _deduplicar(resultados)


def _processar_tabela(tabela: list[list[Any]]) -> list[dict]:
    """
    Processa tabela extraída pelo pdfplumber.
    Tenta múltiplas heurísticas para identificar colunas de ingrediente e quantidade.
    """
    if not tabela or len(tabela) < 2:
        return []

    # Normalizar células None para string vazia
    tabela = [[str(c).strip() if c is not None else "" for c in linha] for linha in tabela]

    PALAVRAS_ING = ["ingrediente", "matéria", "componente", "produto", "alimento",
                    "item", "descrição", "nome"]
    PALAVRAS_QTD = ["quantidade", "qtd", "peso", "gramas", "g)", "massa", "quant", "amount"]

    col_ing = col_qtd = None
    inicio_dados = 1

    for linha_idx in range(min(3, len(tabela))):
        header = [c.lower() for c in tabela[linha_idx]]
        for i, h in enumerate(header):
            if any(p in h for p in PALAVRAS_ING):
                col_ing = i
            if any(p in h for p in PALAVRAS_QTD):
                col_qtd = i
        if col_ing is not None and col_qtd is not None:
            inicio_dados = linha_idx + 1
            break

    # Fallback posicional: coluna com mais números = quantidade; coluna 0 = nome
    if col_ing is None or col_qtd is None:
        if len(tabela[0]) >= 2:
            num_counts = []
            for col in range(len(tabela[0])):
                count = sum(
                    1 for row in tabela[1:]
                    if re.search(r'\d', row[col] if col < len(row) else "")
                )
                num_counts.append(count)
            if num_counts:
                col_qtd  = num_counts.index(max(num_counts))
                col_ing  = 0 if col_qtd != 0 else 1
                inicio_dados = 1

    if col_ing is None or col_qtd is None:
        return []

    resultados = []
    for linha in tabela[inicio_dados:]:
        try:
            if max(col_ing, col_qtd) >= len(linha):
                continue
            nome    = linha[col_ing]
            qtd_str = linha[col_qtd]
            if not nome or not qtd_str:
                continue

            numeros = re.findall(r'\d[\d.,]*', qtd_str)
            if numeros:
                qtd = _num_br(numeros[0])
                nome_limpo = _limpar_nome(nome)
                if qtd > 0 and len(nome_limpo) >= 2 and any(c.isalpha() for c in nome_limpo):
                    resultados.append({"nome": nome_limpo, "quantidade_gramas": qtd})
        except Exception:
            continue

    return resultados


def _deduplicar(resultados: list[dict]) -> list[dict]:
    """Remove duplicatas mantendo a maior quantidade."""
    vistos: dict[str, dict] = {}
    for r in resultados:
        chave = r["nome"].lower().strip()
        if chave not in vistos or r["quantidade_gramas"] > vistos[chave]["quantidade_gramas"]:
            vistos[chave] = r
    return list(vistos.values())


# ═══════════════════════════════════════════════════════════════════════════════
#  PARSER DE COMPOSIÇÃO NUTRICIONAL (fichas técnicas de fornecedores)
# ═══════════════════════════════════════════════════════════════════════════════

# Mapa: sinônimos de rótulo (PT/ES/EN) → campo no banco de dados
_MAPA_NUTRIENTES: dict[str, list[str]] = {
    'energia_kcal':    ['valor calórico', 'valor calorico', 'valor energético',
                        'valor energetico', 'energia kcal', 'calorias', 'energy',
                        'energia', 'caloric value'],
    'energia_kj':      ['energia kj', 'kj', 'kilojoule'],
    'proteina':        ['proteína', 'proteina', 'proteínas', 'proteinas',
                        'protein', 'proteins', 'proteínas totais'],
    'lipideos':        ['gordura total', 'grasa total', 'gorduras totais',
                        'grasas totales', 'lipídeos', 'lipideos', 'lipídios',
                        'gordura', 'grasa', 'fat', 'total fat'],
    'carboidrato':     ['carboidrato', 'carboidratos', 'carbohidrato',
                        'carbohidratos', 'hidratos de carbono',
                        'carbohydrate', 'carbs'],
    'fibra_alimentar': ['fibra alimentar', 'fibra dietética', 'fibra dietetica',
                        'fibra dietaria', 'fibra', 'dietary fiber', 'fiber'],
    'sodio':           ['sódio', 'sodio', 'sodium', 'sal sódico'],
    'colesterol':      ['colesterol', 'cholesterol'],
    'umidade':         ['umidade', 'umidad', 'água', 'agua', 'water', 'moisture'],
    'cinzas':          ['cinzas', 'cenizas', 'ash', 'matéria mineral'],
    'calcio':          ['cálcio', 'calcio', 'calcium'],
    'ferro':           ['ferro', 'hierro', 'iron'],
    'fosforo':         ['fósforo', 'fosforo', 'phosphorus'],
    'potassio':        ['potássio', 'potassio', 'potasio', 'potassium'],
    'magnesio':        ['magnésio', 'magnesio', 'magnesium'],
    'zinco':           ['zinco', 'zinc'],
    'cobre':           ['cobre', 'copper'],
    'manganes':        ['manganês', 'manganes', 'manganese'],
    'retinol':         ['retinol', 'vitamina a', 'vit. a', 'vitamin a'],
    'vitamina_c':      ['vitamina c', 'vit. c', 'ácido ascórbico', 'ascorbic acid'],
    'vitamina_d':      ['vitamina d', 'vit. d', 'vitamin d'],
    'vitamina_e':      ['vitamina e', 'vit. e', 'tocoferol'],
    'vitamina_b12':    ['vitamina b12', 'b12', 'cobalamina', 'cianocobalamina'],
    'tiamina':         ['tiamina', 'vitamina b1', 'b1', 'thiamine'],
    'riboflavina':     ['riboflavina', 'vitamina b2', 'b2'],
    'piridoxina':      ['piridoxina', 'vitamina b6', 'b6'],
    'niacina':         ['niacina', 'vitamina b3', 'b3', 'niacin'],
}


def parse_composicao_nutricional(arquivo_bytes: bytes, extensao: str) -> dict:
    """
    Extrai a composição nutricional de uma ficha técnica de fornecedor.

    Pensado para fichas que mostram VALORES POR 100g de UM ÚNICO ingrediente,
    como as tabelas de fornecedores (ex.: cacao em pó, proteína isolada, etc.).

    Returns:
        dict com campos nutricionais prontos para pré-preencher o formulário
        de Ingredientes de Fornecedor. Inclui 'nome_comercial' se detectado.

    Raises:
        ValueError: se nenhum dado nutricional for extraído.
    """
    ext = extensao.lower().strip(".")
    texto = ""
    pares: list[tuple[str, str]] = []

    if ext == "pdf":
        texto, pares = _extrair_pdf_nutricional(arquivo_bytes)
    elif ext in ("xlsx", "xls"):
        try:
            engine = "openpyxl" if ext == "xlsx" else "xlrd"
            df = pd.read_excel(io.BytesIO(arquivo_bytes), engine=engine, header=None, dtype=str)
            texto = df.to_csv(sep="\t", index=False)
            # Pares: (colA, colB) de cada linha
            for _, row in df.iterrows():
                cells = [str(c).strip() for c in row if pd.notna(c) and str(c).strip()]
                if len(cells) >= 2:
                    pares.append((cells[0], cells[1]))
        except Exception as e:
            raise ValueError(f"Erro ao ler planilha: {e}")
    elif ext in ("csv", "txt"):
        for enc in ("utf-8", "utf-8-sig", "latin-1", "cp1252"):
            try:
                texto = arquivo_bytes.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        else:
            texto = arquivo_bytes.decode("utf-8", errors="replace")
        for linha in texto.splitlines():
            # vírgula só separa colunas quando NÃO é vírgula decimal ("6,5 g")
            partes = re.split(r'[\t;|]|,(?!\d)', linha, maxsplit=1)
            if len(partes) == 2:
                pares.append((partes[0].strip(), partes[1].strip()))
    else:
        raise ValueError(f"Formato '{ext}' não suportado.")

    # Tentar mapear a partir de pares (tabela) — mais preciso
    resultado = _mapear_pares_nutricionais(pares)

    # Complementar com texto livre apenas os campos que a tabela não trouxe
    for campo, valor in _mapear_texto_nutricional(texto).items():
        resultado.setdefault(campo, valor)

    if not any(k != 'nome_comercial' for k in resultado):
        raise ValueError(
            "Não foi possível extrair dados nutricionais da ficha técnica. "
            "Causas prováveis: (1) PDF escaneado sem OCR configurado — instale o Tesseract "
            "(https://github.com/UB-Mannheim/tesseract/wiki) e reinicie o app; "
            "(2) formato de tabela não reconhecido. "
            "Use o formulário manual para cadastrar o ingrediente."
        )

    return resultado


# ─── Configuração do Tesseract (executada uma vez no import) ──────────────────

def _configurar_tesseract():
    """Configura o caminho do Tesseract OCR para Windows se não estiver no PATH."""
    import shutil
    if shutil.which("tesseract"):
        return  # já disponível no PATH
    try:
        import pytesseract
        _PATHS = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        ]
        for p in _PATHS:
            if os.path.exists(p):
                pytesseract.pytesseract.tesseract_cmd = p
                return
    except ImportError:
        pass

_configurar_tesseract()


# ─── Extração de texto de PDFs (3 estratégias em cascata) ─────────────────────

def _extrair_pdf_nutricional(pdf_bytes: bytes) -> tuple[str, list[tuple[str, str]]]:
    """
    Tenta extrair texto e pares (label, valor) de PDF.
    Estratégias: pdfplumber → pymupdf texto → pymupdf OCR (Tesseract).
    """
    texto = ""
    pares: list[tuple[str, str]] = []

    # 1) pdfplumber — funciona bem em PDFs digitais
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for pag in pdf.pages:
                for tabela in (pag.extract_tables() or []):
                    for linha in tabela:
                        if linha and len(linha) >= 2:
                            a = str(linha[0] or "").strip()
                            b = str(linha[1] or "").strip()
                            if a and b:
                                pares.append((a, b))
                t = pag.extract_text(x_tolerance=3, y_tolerance=3) or ""
                texto += t + "\n"
    except Exception:
        pass

    if texto.strip() or pares:
        return texto, pares

    # 2) pymupdf texto — melhor para PDFs com fontes não padrão
    try:
        import fitz
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        for pag in doc:
            texto += pag.get_text("text") + "\n"
        doc.close()
    except Exception:
        pass

    if texto.strip():
        return texto, pares

    # 3) OCR — para PDFs escaneados (imagens)
    texto, pares = _ocr_pdf(pdf_bytes)
    return texto, pares


def _ocr_pdf(pdf_bytes: bytes) -> tuple[str, list[tuple[str, str]]]:
    """
    Renderiza cada página via pymupdf (300 dpi) e aplica Tesseract OCR.
    Combina PSM 4 (colunas) e PSM 6 (bloco) + reconstrução por bounding boxes.
    """
    try:
        import pytesseract
        from PIL import Image, ImageOps
        import fitz
    except ImportError:
        return "", []

    texto_total = ""
    pares: list[tuple[str, str]] = []

    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        for pag in doc:
            mat = fitz.Matrix(300 / 72, 300 / 72)
            pix = pag.get_pixmap(matrix=mat, colorspace=fitz.csGRAY)
            img = Image.frombytes("L", [pix.width, pix.height], pix.samples)
            img = ImageOps.autocontrast(img, cutoff=2)

            # PSM 4: coluna única — bom para tabelas verticais com labels + valores
            txt4 = pytesseract.image_to_string(img, lang="por+eng", config="--psm 4 --oem 3")

            # PSM 6: bounding boxes — reconstrói linhas por posição Y (label+valor juntos)
            linhas_bbox: list[str] = []
            try:
                data = pytesseract.image_to_data(
                    img, lang="por+eng",
                    config="--psm 6 --oem 3",
                    output_type=pytesseract.Output.DICT,
                )
                linhas_pos: dict[int, list[tuple[int, str]]] = {}
                for k in range(len(data["text"])):
                    word = data["text"][k].strip()
                    try:
                        conf = int(data["conf"][k])
                    except (ValueError, TypeError):
                        conf = 0
                    if not word or conf < 20:
                        continue
                    row_key = data["top"][k] // 20
                    linhas_pos.setdefault(row_key, []).append((data["left"][k], word))

                for row_key in sorted(linhas_pos):
                    palavras = sorted(linhas_pos[row_key])
                    linha_txt = " ".join(w for _, w in palavras)
                    linhas_bbox.append(linha_txt)
                    # Extrair par label→valor (letras seguidas de número)
                    m = re.match(
                        r"^([A-Za-zÀ-ÿ][A-Za-zÀ-ÿ ,./\-]+?)\s+([\d][\d,.\-]*).*$",
                        linha_txt,
                    )
                    if m:
                        pares.append((m.group(1).strip(), m.group(2).strip()))

            except Exception:
                # Fallback: split por 2+ espaços em PSM 4
                for linha in txt4.splitlines():
                    linha = linha.strip()
                    if not linha:
                        continue
                    partes = re.split(r"\s{2,}|\t", linha, maxsplit=1)
                    if len(partes) == 2 and partes[0] and partes[1]:
                        pares.append((partes[0].strip(), partes[1].strip()))

            # Bounding box lines primeiro (mais precisas), depois PSM 4 para cobertura
            texto_total += "\n".join(linhas_bbox) + "\n" + txt4 + "\n"

        doc.close()
    except Exception:
        pass

    return texto_total, pares


# ─── Mapeamento de labels nutricionais ────────────────────────────────────────

# Rótulos de sub-nutrientes que NÃO podem ser lidos como o nutriente "pai"
# (ex.: "Gorduras saturadas" não é "Gorduras totais").
_SUBNUTRIENTES = ("saturad", "trans", "monoinsat", "mono-insat", "poliinsat",
                  "poli-insat", "insaturad", "açúcar", "acucar", "azúcar",
                  "sugar", "lactose", "dos quais", "das quais", "of which",
                  "solúvel", "soluvel", "insolúvel", "insoluvel")

# Campos gravados em mg: valores informados em gramas são convertidos.
_CAMPOS_MG = {"sodio", "colesterol", "calcio", "ferro", "fosforo", "potassio",
              "magnesio", "zinco", "cobre", "manganes", "vitamina_c",
              "vitamina_e", "tiamina", "riboflavina", "piridoxina", "niacina"}


def _campo_do_label(label_n: str) -> tuple[str, str] | None:
    """Devolve (campo, sinônimo encontrado) para um rótulo já em minúsculas."""
    if any(sub in label_n for sub in _SUBNUTRIENTES):
        return None
    for campo, sinonimos in _MAPA_NUTRIENTES.items():
        for s in sinonimos:
            # sinônimos curtos ("kj", "b1", "fat") só valem como palavra inteira
            if len(s) <= 3:
                if re.search(rf"(?<![a-zà-ÿ0-9]){re.escape(s)}(?![a-zà-ÿ0-9])", label_n):
                    return campo, s
            elif s in label_n:
                return campo, s
    return None


def _valor_com_unidade(campo: str, label_n: str, valor_txt: str) -> float | None:
    """Primeiro número de valor_txt; converte g → mg quando o campo é em mg."""
    val = _primeiro_numero(valor_txt)
    if val is None:
        return None
    if campo in _CAMPOS_MG:
        unidade = re.search(r"\d\s*(mg|mcg|µg|ug|g)\b", valor_txt.lower())
        em_gramas = (unidade and unidade.group(1) == "g") or (
            not unidade and re.search(r"\(\s*g\s*\)", label_n)
        )
        if em_gramas:
            val *= 1000
    return val


def _registrar_energia(resultado: dict, label_n: str, valor_txt: str) -> None:
    """Lê kcal/kJ de '350 kcal = 1470 kJ' ou de um número sem unidade."""
    kcal, kj = _extrair_kcal_kj(valor_txt)
    if kcal is None and kj is None:
        numero = _primeiro_numero(valor_txt)
        if numero is not None:
            if re.search(r"k\s*j", label_n) and "kcal" not in label_n:
                kj = numero
            else:
                kcal = numero
    if kcal is not None:
        resultado.setdefault("energia_kcal", kcal)
    if kj is not None:
        resultado.setdefault("energia_kj", kj)


def _mapear_pares_nutricionais(pares: list[tuple[str, str]]) -> dict:
    """Mapeia pares (label, valor_texto) para campos do banco."""
    resultado: dict = {}

    for label, valor_txt in pares:
        label_n = label.lower().strip()
        achado = _campo_do_label(label_n)
        if not achado:
            continue
        campo, _ = achado
        if campo in ("energia_kcal", "energia_kj"):
            _registrar_energia(resultado, label_n, valor_txt)
        else:
            val = _valor_com_unidade(campo, label_n, valor_txt)
            if val is not None:
                resultado.setdefault(campo, val)

    return resultado


def _mapear_texto_nutricional(texto: str) -> dict:
    """Extrai valores nutricionais de texto livre linha a linha.
    Suporta labels e valores em linhas separadas (tabelas de duas colunas OCR'd).
    """
    resultado: dict = {}
    linhas = [l.strip() for l in texto.splitlines()]
    n = len(linhas)

    for i, linha in enumerate(linhas):
        linha_n = linha.lower()
        if not linha_n:
            continue

        achado = _campo_do_label(linha_n)
        if not achado:
            continue
        campo, sinonimo = achado
        # O valor é o que vem DEPOIS do nome do nutriente — assim o "12" de
        # "Vitamina B12 2,4 mcg" não é lido como o valor.
        resto = linha[linha_n.index(sinonimo) + len(sinonimo):]

        if campo in ("energia_kcal", "energia_kj"):
            # Juntar esta linha com as próximas até encontrar "X kcal" ou "X kJ"
            janela = resto
            if not re.search(r"\d", janela):
                for j in range(i + 1, min(i + 8, n)):
                    if linhas[j]:
                        janela += " " + linhas[j]
                        if re.search(r"\d\s*k(?:cal|j)", linhas[j], re.IGNORECASE):
                            break
            _registrar_energia(resultado, linha_n, janela)
            continue

        val = _valor_com_unidade(campo, linha_n, resto)
        if val is None:
            # Valor pode estar na linha seguinte ou anterior (tabela 2 colunas)
            for delta in (1, -1, 2, -2):
                j = i + delta
                if 0 <= j < n and linhas[j]:
                    adj = linhas[j]
                    is_label = _campo_do_label(adj.lower()) is not None
                    # Ignorar números de seção ("7.", "7.1 Se...")
                    is_section = bool(re.match(r"^\d+\.\d*\s+[A-Za-z]", adj))
                    if not is_label and not is_section:
                        val = _valor_com_unidade(campo, linha_n, adj)
                        if val is not None:
                            break
        if val is not None:
            resultado.setdefault(campo, val)

    return resultado


def _primeiro_numero(texto: str) -> float | None:
    """Retorna o primeiro número encontrado na string."""
    m = re.search(r'([\d]+[,.][\d]+|[\d]+)', texto)
    if m:
        return float(m.group().replace(',', '.'))
    return None


def _extrair_kcal_kj(texto: str) -> tuple[float | None, float | None]:
    """Extrai kcal e kJ de strings como '235 kcal = 920 kJ'."""
    kcal = kj = None
    m = re.search(r'([\d]+[,.]?[\d]*)\s*kcal', texto, re.IGNORECASE)
    if m:
        kcal = float(m.group(1).replace(',', '.'))
    m = re.search(r'([\d]+[,.]?[\d]*)\s*kj', texto, re.IGNORECASE)
    if m:
        kj = float(m.group(1).replace(',', '.'))
    return kcal, kj
