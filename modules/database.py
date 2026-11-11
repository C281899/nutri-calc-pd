"""
database.py — Operações com SQLite para o NutriCalc P&D
Toda interação com taco.db e nutri_calc.db passa por aqui.
"""

from __future__ import annotations

import sqlite3
import os
import unicodedata
from datetime import datetime
from rapidfuzz import fuzz


def _normalizar(texto: str) -> str:
    """Remove acentos para comparação fuzzy agnóstica a acentuação."""
    return unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii").lower()


# ─── DDL da tabela de densidades ──────────────────────────────────────────────
_SQL_DENSIDADES = """
CREATE TABLE IF NOT EXISTS densidades_ingredientes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nome_ingrediente TEXT NOT NULL UNIQUE,
    densidade_g_ml REAL NOT NULL,
    fonte TEXT DEFAULT 'manual',
    data_cadastro TEXT DEFAULT (datetime('now','localtime'))
);
"""

_DENSIDADES_INICIAIS = [
    ("Leite integral UHT",  1.030, "tabela"),
    ("Leite desnatado",     1.035, "tabela"),
    ("Creme de leite 35%",  1.012, "tabela"),
    ("Iogurte integral",    1.050, "tabela"),
    ("Soro whey",           1.025, "tabela"),
    ("Leite condensado",    1.300, "tabela"),
    ("Óleo de soja",        0.915, "tabela"),
    ("Óleo de palma",       0.910, "tabela"),
    ("Azeite",              0.911, "tabela"),
    ("Óleo de girassol",    0.920, "tabela"),
    ("Água",                1.000, "tabela"),
    ("Suco de laranja",     1.045, "tabela"),
    ("Vinagre",             1.006, "tabela"),
    ("Mel",                 1.420, "tabela"),
    ("Glicerina",           1.261, "tabela"),
    ("Álcool etílico",      0.789, "tabela"),
    ("Xarope de glicose",   1.380, "tabela"),
    ("Extrato de malte",    1.200, "tabela"),
]

# ─── Caminhos ──────────────────────────────────────────────────────────────────
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TACO_DB  = os.path.join(ROOT_DIR, "taco.db")
APP_DB   = os.path.join(ROOT_DIR, "nutri_calc.db")

# ─── Nutrientes (mesma lista do taco_setup.py) ─────────────────────────────────
NUTRIENTES = [
    "umidade", "energia_kcal", "energia_kj", "proteina", "lipideos",
    "colesterol", "carboidrato", "fibra_alimentar", "cinzas",
    "calcio", "magnesio", "manganes", "fosforo", "ferro", "sodio",
    "potassio", "cobre", "zinco", "retinol", "re", "rae",
    "tiamina", "riboflavina", "piridoxina", "niacina",
    "vitamina_c", "vitamina_d", "vitamina_e", "vitamina_b12",
]

NUTRIENTES_EXIBICAO = {
    "energia_kcal":    "Valor Energético (kcal)",
    "energia_kj":      "Valor Energético (kJ)",
    "carboidrato":     "Carboidratos",
    "proteina":        "Proteínas",
    "lipideos":        "Gorduras Totais",
    "fibra_alimentar": "Fibra Alimentar",
    "sodio":           "Sódio",
    "colesterol":      "Colesterol",
    "calcio":          "Cálcio",
    "ferro":           "Ferro",
    "vitamina_c":      "Vitamina C",
}


def _conn_taco() -> sqlite3.Connection:
    if not os.path.exists(TACO_DB):
        raise FileNotFoundError(
            f"Banco TACO não encontrado em '{TACO_DB}'. "
            "Execute 'python data/taco_setup.py' primeiro."
        )
    conn = sqlite3.connect(TACO_DB)
    conn.row_factory = sqlite3.Row
    return conn


def _conn_app() -> sqlite3.Connection:
    if not os.path.exists(APP_DB):
        raise FileNotFoundError(
            f"Banco da aplicação não encontrado em '{APP_DB}'. "
            "Execute 'python data/taco_setup.py' primeiro."
        )
    conn = sqlite3.connect(APP_DB)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def banco_inicializado() -> bool:
    """Verifica se os dois bancos existem e estão populados."""
    return os.path.exists(TACO_DB) and os.path.exists(APP_DB)


# Campos do rótulo que a TACO não fornece e o usuário informa por receita.
# None = não informado.
CAMPOS_EXTRAS_RECEITA = [
    "gordura_saturada_total", "gordura_trans_total",
    "acucares_adicionados_total", "acucares_totais_total",
]

_migrado_para: str | None = None


def migrar_banco(forcar: bool = False):
    """
    Aplica migrações incrementais ao banco da aplicação.
    Idempotente; roda uma vez por processo (o Streamlit reexecuta o script a
    cada clique, então repetir isso seria trabalho jogado fora).
    """
    global _migrado_para
    if not os.path.exists(APP_DB):
        return
    if _migrado_para == APP_DB and not forcar:
        return

    conn = sqlite3.connect(APP_DB)
    try:
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA foreign_keys=ON")

        def _colunas(tabela: str) -> set[str]:
            return {r[1] for r in conn.execute(f"PRAGMA table_info({tabela})")}

        cols_ing = _colunas("receita_ingredientes")
        for col, typ, defval in [
            ("unidade_original",    "TEXT", "'g'"),
            ("quantidade_original", "REAL", "NULL"),
            ("densidade_utilizada", "REAL", "1.0"),
        ]:
            if col not in cols_ing:
                conn.execute(
                    f"ALTER TABLE receita_ingredientes ADD COLUMN {col} {typ} DEFAULT {defval}"
                )

        cols_rec = _colunas("receitas")
        for col in CAMPOS_EXTRAS_RECEITA:
            if col not in cols_rec:
                conn.execute(f"ALTER TABLE receitas ADD COLUMN {col} REAL DEFAULT NULL")

        conn.executescript(_SQL_DENSIDADES)
        count = conn.execute("SELECT COUNT(*) FROM densidades_ingredientes").fetchone()[0]
        if count == 0:
            conn.executemany(
                "INSERT OR IGNORE INTO densidades_ingredientes (nome_ingrediente, densidade_g_ml, fonte) VALUES (?,?,?)",
                _DENSIDADES_INICIAIS,
            )

        conn.executescript("""
            CREATE TABLE IF NOT EXISTS auditoria_triangulacoes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TEXT,
                receita_id INTEGER,
                ingrediente TEXT,
                match_tbca TEXT,
                score_match REAL,
                nutrientes_corrigidos TEXT,
                fonte_por_nutriente TEXT,
                score_confianca REAL,
                nivel_confianca TEXT,
                divergencias TEXT
            );
        """)
        conn.commit()
    finally:
        conn.close()
    _migrado_para = APP_DB


# ─── Conexão TBCA ──────────────────────────────────────────────────────────────

def _conn_tbca() -> sqlite3.Connection | None:
    """Retorna conexão ao TBCA ou None se banco não existir."""
    tbca_path = os.path.join(ROOT_DIR, "tbca.db")
    if not os.path.exists(tbca_path):
        return None
    conn = sqlite3.connect(tbca_path)
    conn.row_factory = sqlite3.Row
    return conn


def get_composicao_tbca_por_id(tbca_id: int) -> dict | None:
    """Retorna composição por 100g de um alimento TBCA pelo id."""
    conn = _conn_tbca()
    if conn is None:
        return None
    try:
        row = conn.execute(
            "SELECT * FROM tbca_alimentos WHERE id=?", (tbca_id,)
        ).fetchone()
        conn.close()
        if row:
            return {k: row[k] for k in row.keys() if k in NUTRIENTES}
        return None
    except Exception:
        return None


# ─── Busca de Ingredientes ─────────────────────────────────────────────────────

_indice_taco: dict = {"mtime": None, "itens": []}


def _carregar_indice_taco() -> list[tuple]:
    """
    Nomes da TACO já normalizados, lidos uma vez por processo (a TACO não muda
    durante o uso). Recarrega sozinho se o arquivo taco.db for recriado.
    """
    mtime = os.path.getmtime(TACO_DB) if os.path.exists(TACO_DB) else None
    if _indice_taco["mtime"] != mtime or not _indice_taco["itens"]:
        conn = _conn_taco()
        try:
            rows = conn.execute(
                "SELECT id, nome_alimento, categoria FROM alimentos_taco"
            ).fetchall()
        finally:
            conn.close()
        _indice_taco["itens"] = [
            (r["id"], r["nome_alimento"], r["categoria"] or "", _normalizar(r["nome_alimento"]))
            for r in rows
        ]
        _indice_taco["mtime"] = mtime
    return _indice_taco["itens"]


def _pontuar(termo_norm: str, nome_norm: str) -> tuple[float, float]:
    """
    (score, proximidade). score decide se o item entra na lista (>= 60);
    proximidade desempata — sem ela "leite integral" e "leite em pó integral"
    empatam em 100% para a busca "leite em pó integral".
    """
    s_tsr = fuzz.token_set_ratio(termo_norm, nome_norm)
    # partial_ratio só ajuda quando o nome é pelo menos tão longo quanto o termo
    s_par = fuzz.partial_ratio(termo_norm, nome_norm) if len(nome_norm) >= len(termo_norm) else 0
    proximidade = fuzz.ratio(termo_norm, nome_norm)
    # termo que é o começo de uma palavra do nome ("ovo" em "ovo de galinha")
    palavras = nome_norm.split()
    if termo_norm in palavras:
        proximidade += 200          # palavra inteira ("sal" em "manteiga com sal")
    elif any(p.startswith(termo_norm) for p in palavras):
        proximidade += 100          # começo de palavra ("sal" em "salmão")
    return max(s_tsr, s_par), proximidade


def buscar_ingrediente(termo: str, limite: int = 10) -> list[dict]:
    """
    Fuzzy matching contra TACO + fornecedores cadastrados.
    Retorna lista de candidatos com: id, nome, fonte, score, categoria,
    ordenada do nome mais parecido para o menos parecido.
    Threshold: 60 (aceita matches razoáveis para autocomplete).
    """
    if not termo or len(termo.strip()) < 2:
        return []

    termo_norm = _normalizar(termo.strip())
    candidatos = []

    try:
        for id_, nome, cat, nome_norm in _carregar_indice_taco():
            score, prox = _pontuar(termo_norm, nome_norm)
            if score >= 60:
                candidatos.append({
                    "id": id_, "nome": nome, "fonte": "TACO",
                    "categoria": cat, "score": score, "_prox": prox,
                })
    except Exception as e:
        print(f"[database] Busca na TACO falhou: {e}")

    try:
        conn = _conn_app()
        try:
            rows = conn.execute(
                "SELECT id, nome_comercial, nome_generico, fabricante FROM ingredientes_fornecedor"
            ).fetchall()
        finally:
            conn.close()

        for r in rows:
            # casa pelo nome comercial ou pelo nome genérico
            melhores = [
                _pontuar(termo_norm, _normalizar(n))
                for n in (r["nome_comercial"], r["nome_generico"]) if n
            ]
            if not melhores:
                continue
            score, prox = max(melhores)
            if score >= 60:
                candidatos.append({
                    "id": r["id"], "nome": r["nome_comercial"], "fonte": "FORNECEDOR",
                    "categoria": f"Fornecedor: {r['fabricante'] or 'N/A'}",
                    "score": score, "_prox": prox,
                })
    except Exception as e:
        print(f"[database] Busca em fornecedores falhou: {e}")

    candidatos.sort(key=lambda x: (x["score"], x["_prox"]), reverse=True)
    for c in candidatos:
        c.pop("_prox", None)
    return candidatos[:limite]


def get_nomes_ingredientes() -> list[str]:
    """Retorna lista de nomes de todos os ingredientes para autocomplete."""
    nomes = []
    try:
        conn = _conn_taco()
        rows = conn.execute("SELECT nome_alimento FROM alimentos_taco ORDER BY nome_alimento").fetchall()
        nomes += [r["nome_alimento"] for r in rows]
        conn.close()
    except Exception:
        pass
    try:
        conn = _conn_app()
        rows = conn.execute("SELECT nome_comercial FROM ingredientes_fornecedor ORDER BY nome_comercial").fetchall()
        nomes += [f"[Fornecedor] {r['nome_comercial']}" for r in rows]
        conn.close()
    except Exception:
        pass
    return sorted(nomes)


# ─── Composição Nutricional ────────────────────────────────────────────────────

def get_composicao_por_100g(fonte: str, id_: int) -> dict | None:
    """
    Retorna dict com todos os nutrientes por 100g do ingrediente.
    fonte: 'TACO' ou 'FORNECEDOR'
    """
    cols = ", ".join(NUTRIENTES)

    if fonte == "TACO":
        try:
            conn = _conn_taco()
            row = conn.execute(
                f"SELECT nome_alimento, fonte, {cols} FROM alimentos_taco WHERE id = ?", (id_,)
            ).fetchone()
            conn.close()
            if row:
                # "fonte_dados" = de onde veio o registro (a tabela alimentos_taco
                # também guarda itens TBCA-USP e USDA) — usado na rastreabilidade.
                return {"nome": row["nome_alimento"],
                        "fonte_dados": row["fonte"] or "TACO 4ª edição",
                        **{n: row[n] for n in NUTRIENTES}}
        except Exception:
            return None

    elif fonte == "FORNECEDOR":
        try:
            conn = _conn_app()
            row = conn.execute(
                f"SELECT nome_comercial, {cols} FROM ingredientes_fornecedor WHERE id = ?", (id_,)
            ).fetchone()
            conn.close()
            if row:
                return {"nome": row["nome_comercial"], "fonte_dados": "Fornecedor",
                        **{n: row[n] for n in NUTRIENTES}}
        except Exception:
            return None

    return None


def resolver_ingrediente(fonte: str, fonte_id, nome: str) -> dict | None:
    """
    Localiza um ingrediente salvo em receita e devolve
    {"fonte", "fonte_id", "nome", "composicao_100g"} ou None.

    Os ids não são estáveis: mudam quando o banco é recriado (reinício do
    servidor, restauração de backup, TACO baixada × base embutida). Por isso o
    id só é aceito se o nome registrado nele for o mesmo da receita; caso
    contrário procura-se pelo nome. Sem isso a receita abriria com a
    composição de OUTRO alimento.
    """
    nome_norm = _normalizar(nome or "")

    if fonte_id is not None:
        comp = get_composicao_por_100g(fonte, fonte_id)
        if comp and _normalizar(comp.get("nome") or "") == nome_norm:
            return {"fonte": fonte, "fonte_id": fonte_id,
                    "nome": comp["nome"], "composicao_100g": comp}

    achado = get_composicao_por_nome(nome, preferir=fonte)
    if achado:
        comp = get_composicao_por_100g(achado["fonte"], achado["id"])
        if comp:
            return {"fonte": achado["fonte"], "fonte_id": achado["id"],
                    "nome": achado["nome"], "composicao_100g": comp}
    return None


def get_composicao_por_nome(nome: str, preferir: str = "TACO") -> dict | None:
    """
    Busca composição pelo nome (sem diferenciar acentos/maiúsculas).
    Procura primeiro na fonte indicada em `preferir`, depois na outra.
    """
    alvo = _normalizar(nome or "")
    if not alvo:
        return None
    cols = ", ".join(NUTRIENTES)

    def _na_taco():
        try:
            id_ = next((i for i, _, _, n in _carregar_indice_taco() if n == alvo), None)
            if id_ is None:
                return None
            conn = _conn_taco()
            try:
                row = conn.execute(
                    f"SELECT id, nome_alimento, {cols} FROM alimentos_taco WHERE id = ?", (id_,)
                ).fetchone()
            finally:
                conn.close()
            if row:
                return {"id": row["id"], "fonte": "TACO", "nome": row["nome_alimento"],
                        **{n: row[n] for n in NUTRIENTES}}
        except Exception:
            pass
        return None

    def _no_fornecedor():
        try:
            conn = _conn_app()
            try:
                rows = conn.execute(
                    f"SELECT id, nome_comercial, {cols} FROM ingredientes_fornecedor ORDER BY id"
                ).fetchall()
            finally:
                conn.close()
            for row in rows:
                if _normalizar(row["nome_comercial"]) == alvo:
                    return {"id": row["id"], "fonte": "FORNECEDOR", "nome": row["nome_comercial"],
                            **{n: row[n] for n in NUTRIENTES}}
        except Exception:
            pass
        return None

    ordem = (_no_fornecedor, _na_taco) if preferir == "FORNECEDOR" else (_na_taco, _no_fornecedor)
    for buscar in ordem:
        achado = buscar()
        if achado:
            return achado
    return None


# ─── Receitas ─────────────────────────────────────────────────────────────────

def salvar_receita(
    nome: str,
    porcao_gramas: float,
    ingredientes: list[dict],
    num_porcoes: int | None = None,
    medida_caseira: str | None = None,
    observacoes: str | None = None,
    receita_id_existente: int | None = None,
    extras: dict | None = None,
) -> int:
    """
    Persiste ou atualiza receita e seus ingredientes.
    ingredientes: [{"nome": str, "fonte": str, "fonte_id": int, "quantidade_gramas": float}, ...]
    extras: valores de CAMPOS_EXTRAS_RECEITA (gordura saturada/trans e açúcares
            informados para a receita inteira). Ausente = não informado.
    Retorna: id da receita salva.
    """
    migrar_banco()
    extras = extras or {}
    vals_extras = [extras.get(c) for c in CAMPOS_EXTRAS_RECEITA]
    conn = _conn_app()
    try:
        rid = None
        if receita_id_existente:
            sets_extras = "".join(f", {c}=?" for c in CAMPOS_EXTRAS_RECEITA)
            cur = conn.execute(
                f"""UPDATE receitas SET nome_produto=?, porcao_gramas=?, num_porcoes=?,
                   medida_caseira=?, observacoes=?, data_atualizacao=?{sets_extras}
                   WHERE id=?""",
                (nome, porcao_gramas, num_porcoes, medida_caseira, observacoes,
                 datetime.now().isoformat(), *vals_extras, receita_id_existente)
            )
            # rowcount 0 = a receita foi excluída enquanto era editada → salva como nova
            if cur.rowcount:
                conn.execute("DELETE FROM receita_ingredientes WHERE receita_id=?", (receita_id_existente,))
                rid = receita_id_existente
        if rid is None:
            cols_extras = "".join(f", {c}" for c in CAMPOS_EXTRAS_RECEITA)
            ph_extras   = ",?" * len(CAMPOS_EXTRAS_RECEITA)
            cur = conn.execute(
                f"INSERT INTO receitas (nome_produto, porcao_gramas, num_porcoes, medida_caseira, observacoes{cols_extras}) VALUES (?,?,?,?,?{ph_extras})",
                (nome, porcao_gramas, num_porcoes, medida_caseira, observacoes, *vals_extras)
            )
            rid = cur.lastrowid

        for i, ing in enumerate(ingredientes):
            conn.execute(
                """INSERT INTO receita_ingredientes
                   (receita_id, nome_ingrediente, fonte, fonte_id, quantidade_gramas,
                    eh_subrecita, subrecita_id, ordem,
                    unidade_original, quantidade_original, densidade_utilizada)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                (rid, ing.get("nome") or ing.get("nome_ingrediente"),
                 ing["fonte"], ing.get("fonte_id"),
                 ing["quantidade_gramas"], ing.get("eh_subrecita", 0),
                 ing.get("subrecita_id"), i,
                 ing.get("unidade_original", "g"),
                 ing.get("quantidade_original"),
                 ing.get("densidade_utilizada", 1.0))
            )

        conn.commit()
        return rid
    except Exception as e:
        conn.rollback()
        raise RuntimeError(f"Erro ao salvar receita: {e}")
    finally:
        conn.close()


def listar_receitas() -> list[dict]:
    """Lista todas as receitas salvas (resumo)."""
    conn = _conn_app()
    rows = conn.execute(
        """SELECT r.id, r.nome_produto, r.porcao_gramas, r.num_porcoes,
                  r.data_criacao, r.data_atualizacao,
                  COUNT(ri.id) as num_ingredientes
           FROM receitas r
           LEFT JOIN receita_ingredientes ri ON ri.receita_id = r.id
           GROUP BY r.id
           ORDER BY r.data_criacao DESC"""
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_receita_completa(receita_id: int) -> dict | None:
    """Retorna receita com todos os ingredientes."""
    conn = _conn_app()
    r = conn.execute("SELECT * FROM receitas WHERE id=?", (receita_id,)).fetchone()
    if not r:
        conn.close()
        return None

    ings = conn.execute(
        "SELECT * FROM receita_ingredientes WHERE receita_id=? ORDER BY ordem",
        (receita_id,)
    ).fetchall()
    conn.close()

    return {**dict(r), "ingredientes": [dict(i) for i in ings]}


def deletar_receita(receita_id: int):
    """Remove receita e seus ingredientes (CASCADE)."""
    conn = _conn_app()
    conn.execute("DELETE FROM receitas WHERE id=?", (receita_id,))
    conn.commit()
    conn.close()


def duplicar_receita(receita_id: int, novo_nome: str) -> int:
    """Cria cópia da receita com novo nome."""
    receita = get_receita_completa(receita_id)
    if not receita:
        raise ValueError("Receita não encontrada.")
    return salvar_receita(
        nome=novo_nome,
        porcao_gramas=receita["porcao_gramas"],
        ingredientes=receita["ingredientes"],
        num_porcoes=receita["num_porcoes"],
        medida_caseira=receita.get("medida_caseira"),
        observacoes=receita.get("observacoes"),
        extras={c: receita.get(c) for c in CAMPOS_EXTRAS_RECEITA},
    )


# ─── Ingredientes de Fornecedor ───────────────────────────────────────────────

def salvar_ingrediente_fornecedor(dados: dict) -> int:
    """Cadastra novo ingrediente industrializado."""
    conn = _conn_app()
    cols = ["nome_comercial", "nome_generico", "fabricante", "observacoes"] + NUTRIENTES
    vals = [dados.get(c) for c in cols]
    placeholders = ", ".join(["?" for _ in cols])
    col_str = ", ".join(cols)
    try:
        cur = conn.execute(
            f"INSERT INTO ingredientes_fornecedor ({col_str}) VALUES ({placeholders})", vals
        )
        conn.commit()
        return cur.lastrowid
    finally:
        conn.close()


def atualizar_ingrediente_fornecedor(id_: int, dados: dict):
    """Atualiza ingrediente de fornecedor existente."""
    conn = _conn_app()
    cols = ["nome_comercial", "nome_generico", "fabricante", "observacoes"] + NUTRIENTES
    sets = ", ".join([f"{c}=?" for c in cols])
    vals = [dados.get(c) for c in cols] + [id_]
    conn.execute(f"UPDATE ingredientes_fornecedor SET {sets} WHERE id=?", vals)
    conn.commit()
    conn.close()


def listar_ingredientes_fornecedor() -> list[dict]:
    """Lista todos os ingredientes de fornecedores cadastrados."""
    conn = _conn_app()
    rows = conn.execute(
        "SELECT * FROM ingredientes_fornecedor ORDER BY nome_comercial"
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def deletar_ingrediente_fornecedor(id_: int):
    """Remove ingrediente de fornecedor."""
    conn = _conn_app()
    conn.execute("DELETE FROM ingredientes_fornecedor WHERE id=?", (id_,))
    conn.commit()
    conn.close()


def get_ingrediente_fornecedor(id_: int) -> dict | None:
    """Retorna ingrediente de fornecedor por ID."""
    conn = _conn_app()
    row = conn.execute(
        "SELECT * FROM ingredientes_fornecedor WHERE id=?", (id_,)
    ).fetchone()
    conn.close()
    return dict(row) if row else None


# ─── Densidades ────────────────────────────────────────────────────────────────

def buscar_densidade_detalhe(nome: str) -> tuple[str, float] | None:
    """
    Procura a densidade (g/mL) de um ingrediente na tabela de densidades.
    Devolve (nome na tabela, densidade) ou None.

    Usa token_sort_ratio com corte alto: "suco de uva" não deve herdar em
    silêncio a densidade de "suco de laranja". Aceita também quando o nome da
    tabela está contido no do ingrediente ("mel" em "mel de abelha").
    """
    if not nome or not os.path.exists(APP_DB):
        return None
    try:
        conn = sqlite3.connect(APP_DB)
        try:
            rows = conn.execute(
                "SELECT nome_ingrediente, densidade_g_ml FROM densidades_ingredientes"
            ).fetchall()
        finally:
            conn.close()
    except sqlite3.Error:
        return None

    nome_norm = _normalizar(nome)
    palavras  = nome_norm.split()
    melhor = None
    for nome_ing, dens in rows:
        alvo  = _normalizar(nome_ing)
        score = fuzz.token_sort_ratio(nome_norm, alvo)
        p_alvo = alvo.split()
        # nome da tabela contido no do ingrediente e começando igual
        # ("óleo de soja" em "óleo de soja refinado")
        if p_alvo and palavras and palavras[0] == p_alvo[0] and set(p_alvo) <= set(palavras):
            score = max(score, 90)
        if melhor is None or score > melhor[0]:
            melhor = (score, nome_ing, dens)
    if melhor and melhor[0] >= 85:
        return melhor[1], melhor[2]
    return None


def buscar_densidade(nome: str) -> float | None:
    """Densidade (g/mL) do ingrediente, ou None se não houver na tabela."""
    achado = buscar_densidade_detalhe(nome)
    return achado[1] if achado else None


def salvar_densidade(nome: str, densidade: float, fonte: str = "manual"):
    """Insere ou atualiza a densidade de um ingrediente na tabela."""
    conn = _conn_app()
    conn.execute(
        """INSERT INTO densidades_ingredientes (nome_ingrediente, densidade_g_ml, fonte)
           VALUES (?,?,?)
           ON CONFLICT(nome_ingrediente) DO UPDATE SET densidade_g_ml=excluded.densidade_g_ml,
                                                       fonte=excluded.fonte""",
        (nome, densidade, fonte),
    )
    conn.commit()
    conn.close()


# ─── Backup / Restauração ─────────────────────────────────────────────────────

def export_backup() -> bytes:
    """
    Exporta todas as receitas e ingredientes de fornecedor como JSON.
    Retorna bytes prontos para st.download_button().
    """
    import json
    from datetime import datetime

    backup: dict = {
        "versao": "1.1",
        "timestamp": datetime.now().isoformat(),
        "receitas": [],
        "ingredientes_fornecedor": [],
    }

    conn = _conn_app()

    receitas = conn.execute("SELECT * FROM receitas ORDER BY id").fetchall()
    for r in receitas:
        rec = dict(r)
        ings = conn.execute(
            "SELECT * FROM receita_ingredientes WHERE receita_id=? ORDER BY ordem",
            (r["id"],),
        ).fetchall()
        rec["ingredientes"] = [dict(i) for i in ings]
        backup["receitas"].append(rec)

    forn = conn.execute(
        "SELECT * FROM ingredientes_fornecedor ORDER BY nome_comercial"
    ).fetchall()
    backup["ingredientes_fornecedor"] = [dict(f) for f in forn]

    conn.close()
    return json.dumps(backup, ensure_ascii=False, indent=2, default=str).encode("utf-8")


def import_backup(json_bytes: bytes) -> dict:
    """
    Restaura receitas e ingredientes de fornecedor a partir de um backup JSON.

    - Importar o mesmo arquivo duas vezes não duplica nada.
    - Os ids mudam na restauração; as receitas são religadas aos ingredientes
      de fornecedor pelo id novo (e, na TACO, pelo nome).

    Retorna contadores: receitas / ingredientes importados e *_ignorados
    (já existentes). Lança ValueError para arquivos malformados.
    """
    import json

    try:
        data = json.loads(json_bytes.decode("utf-8-sig"))
    except Exception:
        raise ValueError("Arquivo inválido — não é um JSON legível.")

    if not isinstance(data, dict):
        raise ValueError("Arquivo inválido — esperado um objeto JSON na raiz.")
    if "receitas" not in data or "ingredientes_fornecedor" not in data:
        raise ValueError(
            "Arquivo inválido — chaves 'receitas' e 'ingredientes_fornecedor' ausentes."
        )
    if not isinstance(data["receitas"], list) or not isinstance(data["ingredientes_fornecedor"], list):
        raise ValueError("Arquivo inválido — 'receitas' e 'ingredientes_fornecedor' devem ser listas.")

    migrar_banco()
    conn = _conn_app()
    cont = {"receitas": 0, "ingredientes": 0,
            "receitas_ignoradas": 0, "ingredientes_ignorados": 0}

    def _chave_forn(nome, fabricante) -> tuple:
        return (_normalizar(nome or ""), _normalizar(fabricante or ""))

    try:
        # ── Ingredientes de fornecedor ──────────────────────────────────────
        cols_forn = ["nome_comercial", "nome_generico", "fabricante", "observacoes"] + NUTRIENTES
        ph_forn   = ", ".join(["?" for _ in cols_forn])
        col_str_forn = ", ".join(cols_forn)

        existentes = {
            _chave_forn(r["nome_comercial"], r["fabricante"]): r["id"]
            for r in conn.execute("SELECT id, nome_comercial, fabricante FROM ingredientes_fornecedor")
        }
        mapa_id_forn: dict = {}   # id no backup → id neste banco

        for item in data["ingredientes_fornecedor"]:
            if not isinstance(item, dict) or not item.get("nome_comercial"):
                continue
            chave = _chave_forn(item.get("nome_comercial"), item.get("fabricante"))
            if chave in existentes:
                novo_id = existentes[chave]
                cont["ingredientes_ignorados"] += 1
            else:
                cur = conn.execute(
                    f"INSERT INTO ingredientes_fornecedor ({col_str_forn}) VALUES ({ph_forn})",
                    [item.get(c) for c in cols_forn],
                )
                novo_id = cur.lastrowid
                existentes[chave] = novo_id
                cont["ingredientes"] += 1
            if item.get("id") is not None:
                mapa_id_forn[item["id"]] = novo_id

        # ── Receitas ────────────────────────────────────────────────────────
        def _assinatura(nome, porcao, ings) -> tuple:
            return (
                _normalizar(nome or ""), round(float(porcao or 0), 4),
                tuple((_normalizar(n or ""), round(float(q or 0), 4)) for n, q in ings),
            )

        ja_salvas = set()
        for r in conn.execute("SELECT id, nome_produto, porcao_gramas FROM receitas").fetchall():
            ings = conn.execute(
                "SELECT nome_ingrediente, quantidade_gramas FROM receita_ingredientes "
                "WHERE receita_id=? ORDER BY ordem, id", (r["id"],)
            ).fetchall()
            ja_salvas.add(_assinatura(r["nome_produto"], r["porcao_gramas"],
                                      [(i[0], i[1]) for i in ings]))

        cols_extras = "".join(f", {c}" for c in CAMPOS_EXTRAS_RECEITA)
        ph_extras   = ",?" * len(CAMPOS_EXTRAS_RECEITA)

        for receita in data["receitas"]:
            if not isinstance(receita, dict) or not receita.get("nome_produto"):
                continue
            ings = [i for i in receita.get("ingredientes", [])
                    if isinstance(i, dict) and i.get("nome_ingrediente")]
            ass = _assinatura(receita.get("nome_produto"), receita.get("porcao_gramas", 100),
                              [(i["nome_ingrediente"], i.get("quantidade_gramas")) for i in ings])
            if ass in ja_salvas:
                cont["receitas_ignoradas"] += 1
                continue

            cur = conn.execute(
                f"""INSERT INTO receitas
                   (nome_produto, porcao_gramas, num_porcoes, medida_caseira, observacoes,
                    data_criacao, data_atualizacao{cols_extras})
                   VALUES (?,?,?,?,?,COALESCE(?, datetime('now','localtime')),?{ph_extras})""",
                (
                    receita.get("nome_produto"),
                    receita.get("porcao_gramas", 100),
                    receita.get("num_porcoes"),
                    receita.get("medida_caseira"),
                    receita.get("observacoes"),
                    receita.get("data_criacao"),
                    receita.get("data_atualizacao"),
                    *[receita.get(c) for c in CAMPOS_EXTRAS_RECEITA],
                ),
            )
            novo_rid = cur.lastrowid

            for ordem, ing in enumerate(ings):
                fonte    = ing.get("fonte", "TACO")
                fonte_id = ing.get("fonte_id")
                if fonte == "FORNECEDOR":
                    fonte_id = mapa_id_forn.get(fonte_id)   # None → resolvido pelo nome ao abrir
                conn.execute(
                    """INSERT INTO receita_ingredientes
                       (receita_id, nome_ingrediente, fonte, fonte_id, quantidade_gramas,
                        eh_subrecita, subrecita_id, ordem,
                        unidade_original, quantidade_original, densidade_utilizada)
                       VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                    (
                        novo_rid, ing.get("nome_ingrediente"), fonte, fonte_id,
                        float(ing.get("quantidade_gramas") or 0),
                        ing.get("eh_subrecita", 0),
                        None,   # subrecita_id do backup apontaria para outra receita
                        ing.get("ordem", ordem),
                        ing.get("unidade_original", "g"),
                        ing.get("quantidade_original"),
                        ing.get("densidade_utilizada", 1.0),
                    ),
                )

            ja_salvas.add(ass)
            cont["receitas"] += 1

        conn.commit()
    except Exception as e:
        conn.rollback()
        raise RuntimeError(f"Erro ao importar backup: {e}")
    finally:
        conn.close()

    return cont
