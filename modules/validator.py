"""
validator.py — Motor de Validação e Triangulação TACO × TBCA

Detecta dados inconsistentes nos registros da TACO usando 5 regras
heurísticas e usa a TBCA (USP) como fonte secundária apenas para
preencher lacunas — valores existentes na TACO nunca são substituídos.

Rastreabilidade completa: cada triangulação gera registro de auditoria
persistido em nutri_calc.db → tabela auditoria_triangulacoes.
"""

from __future__ import annotations

import json
import os
import sqlite3
import unicodedata
from datetime import datetime
from typing import Any

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TBCA_DB  = os.path.join(ROOT_DIR, "tbca.db")
APP_DB   = os.path.join(ROOT_DIR, "nutri_calc.db")

_MACROS      = ["proteina", "lipideos", "carboidrato"]
_LIQUIDOS_KW = ["leite", "suco", "agua", "água", "bebida",
                "liquido", "líquido", "soro", "caldo", "isotonic"]

_TODOS_NUTRIENTES = [
    "umidade", "energia_kcal", "energia_kj", "proteina", "lipideos",
    "colesterol", "carboidrato", "fibra_alimentar", "cinzas",
    "calcio", "magnesio", "manganes", "fosforo", "ferro", "sodio",
    "potassio", "cobre", "zinco", "retinol", "re", "rae",
    "tiamina", "riboflavina", "piridoxina", "niacina",
    "vitamina_c", "vitamina_d", "vitamina_e", "vitamina_b12",
]

# Faixas esperadas por categoria (palavra-chave no nome × nutriente → (min, max))
_FAIXAS = {
    "leite":   {"umidade": (78.0, 93.0), "proteina": (2.5, 8.0)},
    "iogurte": {"umidade": (75.0, 92.0), "proteina": (2.5, 8.0)},
    "oleo":    {"lipideos": (85.0, 100.0), "umidade": (0.0, 3.0)},
    "azeite":  {"lipideos": (85.0, 100.0), "umidade": (0.0, 3.0)},
    "manteiga": {"lipideos": (70.0, 100.0)},
    "acucar":  {"carboidrato": (90.0, 100.0)},
    "açúcar":  {"carboidrato": (90.0, 100.0)},
    "mel":     {"carboidrato": (75.0, 90.0)},
}


# ─── Helpers ──────────────────────────────────────────────────────────────────

def _safe(valor: Any, default: float = 0.0) -> float:
    if valor is None:
        return default
    try:
        return float(valor)
    except (ValueError, TypeError):
        return default


def _normalizar(texto: str) -> str:
    """Remove acentos e converte para minúsculas."""
    if not texto:
        return ""
    nfkd = unicodedata.normalize("NFKD", str(texto))
    return "".join(c for c in nfkd if not unicodedata.combining(c)).lower().strip()


def _nivel_confianca(score: float) -> str:
    if score >= 0.90:
        return "Alta confiança"
    if score >= 0.70:
        return "Confiança moderada"
    return "Baixa confiança — revisar manualmente"


# ─── ETAPA 2.1 — Detecção de inconsistências ──────────────────────────────────

def detectar_inconsistencias(nome_ingrediente: str, dados_taco: dict) -> dict:
    """
    Analisa dados nutricionais retornados da TACO e detecta padrões
    inconsistentes com base em 5 regras heurísticas.

    Returns:
        {
          'inconsistente': bool,
          'nutrientes_zerados': list[str],
          'nutrientes_suspeitos': list[str],
          'score_confianca': float,   # 0.0 a 1.0
          'motivo': str
        }
    """
    nome_norm = _normalizar(nome_ingrediente)
    motivos   = []
    zerados   = []
    suspeitos = []

    prot   = _safe(dados_taco.get("proteina",      0))
    lip    = _safe(dados_taco.get("lipideos",       0))
    carb   = _safe(dados_taco.get("carboidrato",    0))
    ener   = _safe(dados_taco.get("energia_kcal",   0))
    umid   = _safe(dados_taco.get("umidade",        0))
    cinzas = _safe(dados_taco.get("cinzas",         0))

    # Coletar zeros em macros
    for campo in _MACROS:
        if _safe(dados_taco.get(campo, 0)) == 0.0:
            zerados.append(campo)

    # REGRA 1 — Todos os macros zerados
    if prot == 0 and lip == 0 and carb == 0:
        motivos.append("Macronutrientes principais zerados")

    # REGRA 2 — Energia incompatível com macros (Atwater: P*4 + C*4 + L*9)
    energia_calc = prot * 4.0 + carb * 4.0 + lip * 9.0
    if ener > 5 and energia_calc > 0:
        div = abs(ener - energia_calc) / ener
        if div > 0.20:
            motivos.append(
                f"Valor energético ({ener:.0f} kcal) incompatível com macros "
                f"(esperado ≈{energia_calc:.0f} kcal, divergência {div:.0%})"
            )
            suspeitos.append("energia_kcal")

    # REGRA 3 — Alimento líquido com umidade < 50%
    eh_liquido = any(kw in nome_norm for kw in _LIQUIDOS_KW)
    if eh_liquido and umid < 50:
        motivos.append(
            f"Umidade ({umid:.1f}%) inconsistente para alimento líquido"
        )
        suspeitos.append("umidade")

    # REGRA 4 — Soma dos componentes > 105g (impossível)
    soma = prot + lip + carb + umid + cinzas
    if soma > 105:
        motivos.append(
            f"Soma de componentes ({soma:.1f} g/100g) ultrapassa 100 g"
        )
        suspeitos.extend(["proteina", "lipideos", "carboidrato"])

    # REGRA 5 — Faixas esperadas por categoria
    for kw, faixas in _FAIXAS.items():
        if kw in nome_norm:
            for nutriente, (minv, maxv) in faixas.items():
                val = _safe(dados_taco.get(nutriente, 0))
                if not (minv <= val <= maxv):
                    if nutriente not in suspeitos:
                        suspeitos.append(nutriente)

    # Deduplica preservando ordem
    zerados   = list(dict.fromkeys(zerados))
    suspeitos = list(dict.fromkeys(suspeitos))
    inconsistente = len(motivos) > 0

    score = 1.0
    if inconsistente:
        score -= 0.25 * len(motivos)
        score -= 0.05 * len(suspeitos)
        score = max(0.0, min(1.0, score))

    return {
        "inconsistente":        inconsistente,
        "nutrientes_zerados":   zerados,
        "nutrientes_suspeitos": suspeitos,
        "score_confianca":      round(score, 2),
        "motivo":               " | ".join(motivos) if motivos else "OK",
    }


# ─── ETAPA 2.2 — Triangulação com TBCA ───────────────────────────────────────

# Similaridade mínima de nome para considerar que TACO e TBCA falam do MESMO
# alimento. Usa token_sort_ratio: "leite em pó desnatado" × "leite desnatado"
# fica abaixo do corte (são alimentos diferentes), o que não acontecia com
# token_set_ratio, que dava 100% sempre que um nome estava contido no outro.
SCORE_MINIMO_MATCH = 0.88

_cache_tbca: dict = {"mtime": None, "rows": [], "nomes": [], "norm": []}


def _carregar_tbca() -> dict:
    """Lê a TBCA uma vez por processo (recarrega se o arquivo mudar)."""
    try:
        mtime = os.path.getmtime(TBCA_DB)
    except OSError:
        return {"rows": [], "nomes": [], "norm": []}
    if _cache_tbca["mtime"] != mtime:
        con = sqlite3.connect(TBCA_DB)
        con.row_factory = sqlite3.Row
        try:
            rows = [dict(r) for r in con.execute("SELECT * FROM tbca_alimentos")]
        finally:
            con.close()
        _cache_tbca.update(
            mtime=mtime,
            rows=rows,
            nomes=[r["nome_alimento"] for r in rows],
            norm=[_normalizar(r["nome_alimento"]) for r in rows],
        )
    return _cache_tbca


def _buscar_tbca(nome: str) -> tuple[dict | None, str, float]:
    """
    Busca o mesmo alimento na TBCA.
    Retorna (dados_dict | None, nome_match, score_0_a_1).
    """
    try:
        from rapidfuzz import fuzz, process
        base = _carregar_tbca()
    except Exception:
        return None, "", 0.0

    if not base["rows"]:
        return None, "", 0.0

    result = process.extractOne(
        _normalizar(nome), base["norm"], scorer=fuzz.token_sort_ratio
    )
    if not result:
        return None, "", 0.0

    _, score_bruto, idx = result
    score = score_bruto / 100.0
    if score >= SCORE_MINIMO_MATCH:
        return dict(base["rows"][idx]), base["nomes"][idx], score
    return None, base["nomes"][idx], score


def _mesmo_alimento(dados_taco: dict, dados_tbca: dict) -> bool:
    """
    Confere se os dois registros são compatíveis entre si: quando ambos têm
    energia ou umidade, os valores não podem estar muito distantes.
    """
    e1, e2 = _safe(dados_taco.get("energia_kcal")), _safe(dados_tbca.get("energia_kcal"))
    if e1 > 0 and e2 > 0 and abs(e1 - e2) / max(e1, e2) > 0.25:
        return False
    u1, u2 = _safe(dados_taco.get("umidade")), _safe(dados_tbca.get("umidade"))
    if u1 > 0 and u2 > 0 and abs(u1 - u2) > 15:
        return False
    return True


def triangular_com_tbca(
    nome_ingrediente: str,
    dados_taco: dict,
    inconsistencias: dict,
) -> dict:
    """
    Procura o mesmo alimento na TBCA e PREENCHE LACUNAS da TACO.

    Regras de segurança:
      - só há triangulação se o nome casar com score >= SCORE_MINIMO_MATCH e os
        dois registros forem compatíveis (energia/umidade próximas);
      - um valor da TACO maior que zero NUNCA é substituído nem promediado;
      - só são preenchidos nutrientes ausentes/zerados na TACO que existam na TBCA;
      - diferenças entre as fontes são apenas relatadas em 'divergencias'.

    Returns:
        {
          'dados_finais': dict,
          'fonte_por_nutriente': dict,        # 'TACO' | 'TBCA'
          'match_tbca': str,
          'score_match': float,
          'triangulacao_aplicada': bool,
          'nutrientes_corrigidos': list[dict],
          'nivel_confianca': str,
          'divergencias': list[str]
        }
    """
    dados_tbca, match_nome, score = _buscar_tbca(nome_ingrediente)
    score_inc = inconsistencias.get("score_confianca", 1.0)

    def _sem_triangulacao(motivo: str) -> dict:
        return {
            "dados_finais":          dados_taco,
            "fonte_por_nutriente":   {n: "TACO" for n in _TODOS_NUTRIENTES},
            "match_tbca":            "",
            "score_match":           round(score, 3),
            "triangulacao_aplicada": False,
            "nutrientes_corrigidos": [],
            "nivel_confianca":       _nivel_confianca(score_inc),
            "divergencias":          [motivo],
        }

    if dados_tbca is None:
        return _sem_triangulacao(
            f"Sem equivalente na TBCA com similaridade ≥ {SCORE_MINIMO_MATCH:.0%}"
        )
    if not _mesmo_alimento(dados_taco, dados_tbca):
        return _sem_triangulacao(
            f"'{match_nome}' (TBCA) tem energia/umidade incompatíveis — não utilizado"
        )

    dados_finais        = dict(dados_taco)
    fonte_por_nutriente = {}
    corrigidos          = []
    divergencias        = []

    for nutriente in _TODOS_NUTRIENTES:
        val_taco = _safe(dados_taco.get(nutriente, 0))
        val_tbca = _safe(dados_tbca.get(nutriente, 0))
        fonte_por_nutriente[nutriente] = "TACO"

        if val_taco == 0 and val_tbca > 0:
            dados_finais[nutriente]        = val_tbca
            fonte_por_nutriente[nutriente] = "TBCA"
            corrigidos.append({"nutriente": nutriente,
                               "val_taco": val_taco, "val_tbca": val_tbca})
        elif val_taco > 0 and val_tbca > 0:
            diff = abs(val_taco - val_tbca) / max(val_taco, val_tbca)
            if diff >= 0.20:
                divergencias.append(
                    f"{nutriente}: TACO={val_taco:.3g} vs TBCA={val_tbca:.3g} "
                    f"(dif. {diff:.0%}) — mantido o valor da TACO"
                )

    if not corrigidos:
        res = _sem_triangulacao("TBCA não tinha nenhum valor ausente na TACO")
        res["match_tbca"]   = match_nome
        res["divergencias"] = divergencias or res["divergencias"]
        return res

    return {
        "dados_finais":          dados_finais,
        "fonte_por_nutriente":   fonte_por_nutriente,
        "match_tbca":            match_nome,
        "score_match":           round(score, 3),
        "triangulacao_aplicada": True,
        "nutrientes_corrigidos": corrigidos,
        "nivel_confianca":       _nivel_confianca(score),
        "divergencias":          divergencias,
    }


# ─── ETAPA 2.3 — Registro de auditoria ───────────────────────────────────────

_SQL_AUDITORIA = """
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
"""


def garantir_tabela_auditoria():
    """Cria tabela de auditoria se não existir (idempotente)."""
    if not os.path.exists(APP_DB):
        return
    try:
        con = sqlite3.connect(APP_DB)
        con.executescript(_SQL_AUDITORIA)
        con.commit()
        con.close()
    except Exception as e:
        print(f"[validator] Erro ao criar tabela de auditoria: {e}")


def registrar_auditoria(
    receita_id: int,
    ingrediente: str,
    resultado: dict,
) -> None:
    """
    Persiste um registro completo de triangulação para rastreabilidade.
    Falha silenciosamente para não interromper o fluxo de cálculo.
    """
    if not os.path.exists(APP_DB):
        return
    try:
        garantir_tabela_auditoria()
        con = sqlite3.connect(APP_DB)
        con.execute("""
            INSERT INTO auditoria_triangulacoes (
                timestamp, receita_id, ingrediente, match_tbca, score_match,
                nutrientes_corrigidos, fonte_por_nutriente, score_confianca,
                nivel_confianca, divergencias
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            datetime.now().isoformat(),
            receita_id,
            ingrediente,
            resultado.get("match_tbca", ""),
            resultado.get("score_match", 0.0),
            json.dumps(resultado.get("nutrientes_corrigidos", []),
                       ensure_ascii=False),
            json.dumps(resultado.get("fonte_por_nutriente", {}),
                       ensure_ascii=False),
            resultado.get("score_match", 0.0),
            resultado.get("nivel_confianca", ""),
            json.dumps(resultado.get("divergencias", []),
                       ensure_ascii=False),
        ))
        con.commit()
        con.close()
    except Exception as e:
        print(f"[validator] Erro ao registrar auditoria: {e}")
