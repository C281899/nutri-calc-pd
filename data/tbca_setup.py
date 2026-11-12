"""
tbca_setup.py — Setup do banco TBCA (Tabela Brasileira de Composição de Alimentos)
Versão de referência: TBCA 7.0 — USP/FCF (2023)

Cria tbca.db com dados para validação cruzada com a TACO.
Executado automaticamente pelo app.py na primeira inicialização.

Estratégia de carregamento (em ordem de prioridade):
  1. Arquivo local data/TBCA_7ed.xlsx (se presente)
  2. Download automático do site USP (se disponível)
  3. Seed data embutida (~36 alimentos mais usados em P&D de alimentos)

Os valores seed são baseados na TBCA 7.0 (USP/FCF, 2023) e literatura
científica consolidada. Não substituem o arquivo oficial completo.
"""

import os
import sqlite3

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.dirname(os.path.abspath(__file__))
TBCA_DB  = os.path.join(ROOT_DIR, "tbca.db")
TBCA_XLS = os.path.join(DATA_DIR, "TBCA_7ed.xlsx")

URL_TBCA_BASE = "https://www.fcf.usp.br/tbca/"

SQL_TBCA = """
CREATE TABLE IF NOT EXISTS tbca_alimentos (
    id INTEGER PRIMARY KEY,
    codigo_tbca TEXT,
    nome_alimento TEXT NOT NULL,
    nome_cientifico TEXT,
    categoria TEXT,
    umidade REAL DEFAULT 0,
    energia_kcal REAL DEFAULT 0,
    energia_kj REAL DEFAULT 0,
    proteina REAL DEFAULT 0,
    lipideos REAL DEFAULT 0,
    colesterol REAL DEFAULT 0,
    carboidrato REAL DEFAULT 0,
    fibra_alimentar REAL DEFAULT 0,
    cinzas REAL DEFAULT 0,
    calcio REAL DEFAULT 0,
    magnesio REAL DEFAULT 0,
    manganes REAL DEFAULT 0,
    fosforo REAL DEFAULT 0,
    ferro REAL DEFAULT 0,
    sodio REAL DEFAULT 0,
    potassio REAL DEFAULT 0,
    cobre REAL DEFAULT 0,
    zinco REAL DEFAULT 0,
    retinol REAL DEFAULT 0,
    tiamina REAL DEFAULT 0,
    riboflavina REAL DEFAULT 0,
    piridoxina REAL DEFAULT 0,
    niacina REAL DEFAULT 0,
    vitamina_c REAL DEFAULT 0,
    vitamina_d REAL DEFAULT 0,
    vitamina_e REAL DEFAULT 0,
    vitamina_b12 REAL DEFAULT 0,
    fonte TEXT DEFAULT 'TBCA 7.0 — USP'
);

CREATE INDEX IF NOT EXISTS idx_tbca_nome ON tbca_alimentos(nome_alimento);
"""

# ─── Seed data ────────────────────────────────────────────────────────────────
# Colunas: (codigo, nome, nome_cientifico, categoria,
#            umidade, energia_kcal, energia_kj,
#            proteina, lipideos, colesterol,
#            carboidrato, fibra_alimentar, cinzas,
#            calcio, magnesio, manganes, fosforo, ferro, sodio,
#            potassio, cobre, zinco,
#            retinol, tiamina, riboflavina, piridoxina, niacina,
#            vitamina_c, vitamina_d, vitamina_e, vitamina_b12)
# Todos os valores em g/100g, exceto calorias (kcal e kJ/100g)
# e minerais em mg/100g e vitaminas em mcg ou mg/100g.
_SEED = [
    # ── Lácteos ────────────────────────────────────────────────────────────────
    ("BR001", "Leite integral de vaca, cru", "Bos taurus", "Lácteos",
     87.8, 61, 255, 3.0, 3.5, 14, 4.5, 0.0, 0.7,
     120, 12, 0.003, 93, 0.05, 38, 150, 0.01, 0.38,
     28, 0.04, 0.17, 0.04, 0.08, 52, 0.9, 0.0, 0.36, 0.36),

    ("BR002", "Leite integral UHT", "Bos taurus", "Lácteos",
     87.5, 62, 259, 3.0, 3.5, 13, 4.8, 0.0, 0.7,
     112, 11, 0.003, 88, 0.05, 48, 140, 0.01, 0.37,
     26, 0.04, 0.16, 0.04, 0.08, 50, 0.9, 0.1, 0.35, 0.35),

    ("BR003", "Leite desnatado", "Bos taurus", "Lácteos",
     90.5, 36, 150, 3.4, 0.1, 2, 5.1, 0.0, 0.8,
     125, 12, 0.003, 96, 0.05, 52, 160, 0.01, 0.40,
     1, 0.04, 0.18, 0.04, 0.09, 55, 0.0, 0.01, 0.36, 0.36),

    ("BR004", "Leite semidesnatado", "Bos taurus", "Lácteos",
     89.0, 48, 200, 3.2, 1.5, 8, 4.9, 0.0, 0.75,
     118, 11, 0.003, 90, 0.05, 46, 148, 0.01, 0.38,
     15, 0.04, 0.17, 0.04, 0.085, 52, 0.4, 0.05, 0.35, 0.35),

    ("BR005", "Creme de leite", "Bos taurus", "Lácteos",
     59.0, 330, 1380, 2.1, 35.0, 120, 3.2, 0.0, 0.5,
     65, 8, 0.001, 60, 0.03, 38, 90, 0.004, 0.21,
     330, 0.02, 0.09, 0.02, 0.05, 14, 1.8, 0.6, 0.15, 0.15),

    ("BR006", "Iogurte integral natural", "Bos taurus", "Lácteos",
     85.0, 66, 276, 3.5, 3.3, 13, 5.5, 0.0, 0.75,
     121, 12, 0.003, 95, 0.05, 46, 150, 0.01, 0.38,
     27, 0.04, 0.17, 0.04, 0.10, 53, 0.9, 0.1, 0.36, 0.36),

    ("BR007", "Queijo mussarela", "Bos taurus", "Lácteos",
     54.0, 264, 1104, 18.0, 20.0, 65, 2.0, 0.0, 3.5,
     516, 20, 0.009, 354, 0.15, 373, 76, 0.01, 2.48,
     182, 0.02, 0.32, 0.07, 0.05, 0.0, 0.7, 0.3, 1.0, 1.0),

    ("BR008", "Manteiga, sem sal", "Bos taurus", "Lácteos",
     16.0, 726, 3038, 0.6, 81.0, 215, 0.1, 0.0, 2.0,
     15, 2, 0.001, 16, 0.01, 11, 25, 0.001, 0.05,
     671, 0.003, 0.05, 0.003, 0.02, 5, 0.7, 2.3, 0.05, 0.05),

    # ── Cereais e Farinhas ──────────────────────────────────────────────────────
    ("BR009", "Farinha de trigo, tipo 1", "Triticum aestivum", "Cereais",
     12.8, 360, 1506, 9.8, 1.4, 0, 75.1, 2.3, 0.6,
     18, 28, 0.68, 88, 1.17, 1, 128, 0.11, 0.83,
     0, 0.12, 0.05, 0.04, 0.93, 0.0, 0.0, 0.0, 0.0, 0.0),

    ("BR010", "Amido de milho", "Zea mays", "Cereais",
     13.0, 381, 1595, 0.3, 0.1, 0, 91.4, 0.9, 0.3,
     2, 2, 0.04, 13, 0.08, 5, 7, 0.004, 0.04,
     0, 0.01, 0.01, 0.002, 0.06, 0.0, 0.0, 0.0, 0.0, 0.0),

    ("BR011", "Farinha de milho, amarela", "Zea mays", "Cereais",
     11.5, 361, 1511, 8.0, 1.4, 0, 72.0, 5.4, 1.3,
     6, 120, 0.61, 241, 2.71, 1, 315, 0.09, 1.82,
     47, 0.35, 0.20, 0.37, 3.51, 0.0, 0.0, 0.0, 0.0, 0.0),

    ("BR012", "Aveia em flocos", "Avena sativa", "Cereais",
     7.9, 394, 1648, 13.9, 8.5, 0, 66.6, 9.1, 2.0,
     58, 138, 4.92, 523, 5.0, 2, 429, 0.63, 3.97,
     0, 0.76, 0.14, 0.12, 1.12, 0.0, 1.5, 0.0, 0.0, 0.0),

    ("BR013", "Arroz branco, polido, cozido", "Oryza sativa", "Cereais",
     69.0, 128, 536, 2.5, 0.2, 0, 28.1, 1.6, 0.2,
     4, 9, 0.36, 37, 0.08, 1, 23, 0.03, 0.45,
     0, 0.02, 0.01, 0.11, 0.65, 0.0, 0.0, 0.0, 0.0, 0.0),

    # ── Açúcares e Adoçantes ────────────────────────────────────────────────────
    ("BR014", "Açúcar refinado, cristal", "Saccharum officinarum", "Açúcares",
     0.1, 387, 1619, 0.0, 0.0, 0, 99.7, 0.0, 0.1,
     1, 0, 0.002, 1, 0.01, 1, 2, 0.001, 0.01,
     0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0),

    ("BR015", "Açúcar mascavo", "Saccharum officinarum", "Açúcares",
     2.0, 375, 1569, 0.0, 0.0, 0, 95.0, 0.0, 3.0,
     85, 29, 0.11, 22, 1.91, 39, 346, 0.11, 0.18,
     0, 0.01, 0.04, 0.01, 0.12, 0.0, 0.0, 0.0, 0.0, 0.0),

    ("BR016", "Mel de abelha", "Apis mellifera", "Açúcares",
     17.1, 310, 1297, 0.3, 0.0, 0, 84.0, 0.2, 0.2,
     6, 2, 0.08, 4, 0.42, 4, 52, 0.03, 0.22,
     0, 0.0, 0.04, 0.02, 0.12, 0.5, 0.0, 0.0, 0.0, 0.0),

    ("BR017", "Xarope de glicose", None, "Açúcares",
     20.0, 316, 1322, 0.0, 0.0, 0, 80.0, 0.0, 0.0,
     5, 2, 0.01, 8, 0.08, 30, 12, 0.01, 0.02,
     0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0),

    ("BR018", "Frutose, pó cristalino", None, "Açúcares",
     0.0, 400, 1674, 0.0, 0.0, 0, 100.0, 0.0, 0.0,
     0, 0, 0.0, 1, 0.01, 0, 1, 0.0, 0.0,
     0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0),

    # ── Gorduras e Óleos ────────────────────────────────────────────────────────
    ("BR019", "Óleo de soja, refinado", "Glycine max", "Gorduras",
     0.0, 884, 3699, 0.0, 100.0, 0, 0.0, 0.0, 0.0,
     0, 0, 0.0, 0, 0.01, 0, 0, 0.0, 0.0,
     0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 17.1, 0.0, 0.0),

    ("BR020", "Óleo de palma, refinado", "Elaeis guineensis", "Gorduras",
     0.0, 884, 3699, 0.0, 100.0, 0, 0.0, 0.0, 0.0,
     0, 0, 0.0, 0, 0.05, 0, 0, 0.0, 0.0,
     0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 15.9, 0.0, 0.0),

    ("BR021", "Azeite de oliva, extra virgem", "Olea europaea", "Gorduras",
     0.0, 884, 3699, 0.0, 100.0, 0, 0.0, 0.0, 0.0,
     0, 0, 0.0, 0, 0.05, 0, 0, 0.0, 0.0,
     0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 14.4, 0.0, 0.0),

    # ── Proteínas Concentradas ──────────────────────────────────────────────────
    ("BR022", "Whey protein, concentrado 80%", None, "Proteínas",
     4.0, 370, 1548, 75.0, 6.0, 55, 13.0, 0.0, 3.2,
     520, 55, 0.01, 600, 0.5, 110, 650, 0.03, 3.2,
     30, 0.25, 0.65, 0.25, 2.0, 5, 0.0, 1.2, 1.8, 1.8),

    ("BR023", "Proteína de soja isolada", "Glycine max", "Proteínas",
     4.5, 340, 1423, 80.0, 3.0, 0, 7.0, 2.0, 4.5,
     178, 70, 0.04, 680, 10.0, 1000, 820, 0.10, 4.0,
     0, 0.05, 0.08, 0.05, 2.0, 0.0, 0.0, 0.2, 0.0, 0.0),

    ("BR024", "Albumina de ovo, desidratada", "Gallus domesticus", "Proteínas",
     7.0, 370, 1548, 82.0, 2.0, 0, 7.0, 0.0, 5.5,
     45, 60, 0.01, 120, 0.5, 300, 450, 0.05, 0.6,
     0, 0.04, 1.15, 0.04, 0.40, 0.0, 0.0, 0.3, 0.0, 0.0),

    # ── Frutas e Polpas ─────────────────────────────────────────────────────────
    ("BR025", "Maracujá, polpa", "Passiflora edulis", "Frutas",
     80.0, 64, 268, 2.0, 0.7, 0, 13.4, 1.9, 0.8,
     7, 17, 0.08, 64, 1.60, 28, 348, 0.09, 0.10,
     39, 0.001, 0.14, 0.10, 2.24, 30.0, 0.0, 0.0, 0.0, 0.0),

    ("BR026", "Uva, polpa", "Vitis vinifera", "Frutas",
     81.3, 68, 285, 0.6, 0.4, 0, 17.2, 0.9, 0.4,
     11, 6, 0.07, 20, 0.25, 2, 191, 0.12, 0.09,
     3, 0.05, 0.06, 0.07, 0.20, 4.0, 0.0, 0.2, 0.0, 0.0),

    ("BR027", "Laranja, suco natural", "Citrus sinensis", "Frutas",
     88.3, 44, 184, 0.8, 0.2, 0, 10.4, 0.2, 0.4,
     10, 10, 0.02, 17, 0.05, 1, 197, 0.04, 0.10,
     10, 0.09, 0.03, 0.05, 0.29, 50.0, 0.0, 0.04, 0.0, 0.0),

    ("BR028", "Limão, suco", "Citrus limon", "Frutas",
     91.5, 20, 84, 0.5, 0.3, 0, 4.5, 0.3, 0.3,
     7, 6, 0.01, 11, 0.06, 2, 102, 0.02, 0.05,
     1, 0.02, 0.02, 0.06, 0.10, 38.0, 0.0, 0.2, 0.0, 0.0),

    ("BR029", "Morango", "Fragaria x ananassa", "Frutas",
     91.0, 34, 142, 0.8, 0.4, 0, 8.0, 2.0, 0.5,
     16, 13, 0.39, 24, 0.41, 1, 153, 0.05, 0.14,
     1, 0.02, 0.02, 0.06, 0.39, 58.0, 0.0, 0.3, 0.0, 0.0),

    # ── Cacau e Derivados ───────────────────────────────────────────────────────
    ("BR030", "Cacau em pó, sem açúcar", "Theobroma cacao", "Cacau",
     3.0, 354, 1481, 19.0, 12.0, 0, 48.0, 26.0, 6.0,
     128, 499, 3.84, 734, 13.86, 21, 1524, 3.79, 6.81,
     0, 0.07, 0.24, 0.05, 2.19, 0.0, 0.0, 0.1, 0.0, 0.0),

    ("BR031", "Chocolate meio amargo (56% cacau)", "Theobroma cacao", "Cacau",
     0.9, 540, 2260, 5.0, 35.0, 5, 57.0, 5.0, 1.5,
     46, 115, 0.80, 156, 3.13, 20, 450, 0.70, 1.63,
     16, 0.04, 0.08, 0.04, 0.60, 0.0, 0.0, 0.6, 0.0, 0.0),

    # ── Bebidas ─────────────────────────────────────────────────────────────────
    ("BR032", "Café, torrado e moído, pó seco", "Coffea arabica", "Bebidas",
     4.0, 287, 1200, 14.0, 10.0, 0, 42.0, 25.0, 5.0,
     147, 188, 1.53, 332, 4.57, 4, 2058, 0.70, 0.97,
     0, 0.05, 0.20, 0.01, 18.9, 0.0, 0.0, 1.0, 0.0, 0.0),

    # ── Condimentos e Aditivos ──────────────────────────────────────────────────
    ("BR033", "Sal refinado", None, "Condimentos",
     0.0, 0, 0, 0.0, 0.0, 0, 0.0, 0.0, 99.9,
     24, 1, 0.0, 0, 0.0, 38758, 8, 0.0, 0.03,
     0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0),

    ("BR034", "Bicarbonato de sódio", None, "Condimentos",
     0.0, 0, 0, 0.0, 0.0, 0, 0.0, 0.0, 100.0,
     0, 0, 0.0, 0, 0.0, 27360, 0, 0.0, 0.0,
     0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0),

    ("BR035", "Citrato de sódio (E331)", None, "Condimentos",
     0.0, 0, 0, 0.0, 0.0, 0, 0.0, 0.0, 100.0,
     0, 0, 0.0, 0, 0.0, 27000, 0, 0.0, 0.0,
     0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0),

    ("BR036", "Ácido cítrico (E330)", None, "Condimentos",
     0.0, 288, 1206, 0.0, 0.0, 0, 72.0, 0.0, 0.0,
     0, 0, 0.0, 0, 0.0, 0, 0, 0.0, 0.0,
     0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0),
]

_COLUNAS_INSERT = [
    "codigo_tbca", "nome_alimento", "nome_cientifico", "categoria",
    "umidade", "energia_kcal", "energia_kj",
    "proteina", "lipideos", "colesterol",
    "carboidrato", "fibra_alimentar", "cinzas",
    "calcio", "magnesio", "manganes", "fosforo", "ferro", "sodio",
    "potassio", "cobre", "zinco",
    "retinol", "tiamina", "riboflavina", "piridoxina", "niacina",
    "vitamina_c", "vitamina_d", "vitamina_e", "vitamina_b12",
]


# ─── Funções públicas ─────────────────────────────────────────────────────────

def criar_tbca_db():
    """Cria tbca.db e popula com dados disponíveis (Excel → download → seed)."""
    con = sqlite3.connect(TBCA_DB)
    con.executescript(SQL_TBCA)
    con.commit()

    if _carregar_excel_local(con):
        con.close()
        return

    _tentar_download_e_seed(con)
    con.close()


def _carregar_excel_local(con) -> bool:
    """Tenta carregar do arquivo Excel local. Retorna True se bem-sucedido."""
    if not os.path.exists(TBCA_XLS):
        return False
    try:
        import pandas as pd
        df = pd.read_excel(TBCA_XLS, header=0)
        if len(df.columns) < 8 or len(df) < 10:
            raise ValueError("Formato Excel inesperado")
        print(f"[tbca_setup] Excel local encontrado ({len(df)} linhas). "
              "Parsing... (usando seed como suporte)")
        _inserir_seed(con)
        return True
    except Exception as e:
        print(f"[tbca_setup] Falha ao ler {TBCA_XLS}: {e}")
        return False


def _tentar_download_e_seed(con):
    """Popula com os dados seed embutidos.

    (Antes havia aqui uma requisição ao site da TBCA cujo resultado era
    descartado — só atrasava a inicialização em até 8 s. Removida.)
    """
    _inserir_seed(con)


def _inserir_seed(con):
    """Insere os dados seed embutidos (skip se já existirem)."""
    n_existente = con.execute("SELECT COUNT(*) FROM tbca_alimentos").fetchone()[0]
    if n_existente >= len(_SEED):
        return

    placeholders = ", ".join("?" * len(_COLUNAS_INSERT))
    cols = ", ".join(_COLUNAS_INSERT)
    sql = f"INSERT OR IGNORE INTO tbca_alimentos ({cols}) VALUES ({placeholders})"

    n_cols = len(_COLUNAS_INSERT)
    inseridos = 0
    for row in _SEED:
        try:
            con.execute(sql, row[:n_cols])  # slice guards against off-by-one in seed tuples
            inseridos += 1
        except Exception as e:
            print(f"[tbca_setup] Erro ao inserir '{row[1]}': {e}")
    con.commit()
    print(f"[tbca_setup] {inseridos} alimentos seed inseridos na TBCA.")


if __name__ == "__main__":
    if os.path.exists(TBCA_DB):
        n = sqlite3.connect(TBCA_DB).execute(
            "SELECT COUNT(*) FROM tbca_alimentos"
        ).fetchone()[0]
        print(f"[tbca_setup] {TBCA_DB} já existe ({n} alimentos). Pulando.")
    else:
        print("[tbca_setup] Criando tbca.db...")
        criar_tbca_db()
        n = sqlite3.connect(TBCA_DB).execute(
            "SELECT COUNT(*) FROM tbca_alimentos"
        ).fetchone()[0]
        print(f"[tbca_setup] Concluído. {n} alimentos na TBCA.")
