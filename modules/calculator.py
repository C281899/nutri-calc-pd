"""
calculator.py — Motor de Cálculo Nutricional
Implementa toda a lógica de cálculo conforme RDC 429/2020 (ANVISA).
calcular_composicao_receita aceita: ingredientes, receita_id (opcional), validar_taco (opcional).
"""

from __future__ import annotations
from typing import Any

# ─── Valores Diários de Referência (IN 75/2020, Anexo II) ─────────────────────
VALORES_DIARIOS = {
    "energia_kcal":         2000.0,
    "energia_kj":           8400.0,
    "carboidrato":           300.0,
    "acucares_adicionados":   50.0,
    "proteina":               50.0,
    "lipideos":               65.0,
    "gordura_saturada":       20.0,
    "gordura_trans":           2.0,
    "fibra_alimentar":        25.0,
    "sodio":                2000.0,
}

KCAL_PARA_KJ = 4.184

# Nutrientes que não têm %VD estabelecido
SEM_VD = {"acucares_totais", "colesterol"}

# ─── Quantidades não significativas (IN 75/2020, Anexo IV) ────────────────────
# Declaradas como zero quando o valor em 100 g E na porção não ultrapassa o limite.
LIMITES_NAO_SIGNIFICATIVOS = {
    "energia_kcal":         4.0,    # kcal
    "carboidrato":          0.5,    # g
    "acucares_totais":      0.5,
    "proteina":             0.5,
    "lipideos":             0.5,
    "gordura_saturada":     0.1,
    "gordura_trans":        0.1,
    "fibra_alimentar":      0.5,
    "sodio":                5.0,    # mg
    "colesterol":           5.0,    # mg
}

# Unidade de cada linha do rótulo (define a regra de arredondamento abaixo de 1)
UNIDADES_ROTULO = {
    "energia_kcal": "kcal", "energia_kj": "kJ",
    "carboidrato": "g", "acucares_totais": "g", "acucares_adicionados": "g",
    "proteina": "g", "lipideos": "g", "gordura_saturada": "g",
    "gordura_trans": "g", "fibra_alimentar": "g",
    "sodio": "mg", "colesterol": "mg",
}


def _safe_float(valor: Any, default: float = 0.0) -> float:
    """Converte valor para float, retornando default se None/inválido."""
    if valor is None:
        return default
    try:
        return float(valor)
    except (ValueError, TypeError):
        return default


def _half_up(valor: float, casas: int = 0) -> float:
    """Arredonda "meio para cima" (0,25 → 0,3), como exige a IN 75/2020.
    O round() do Python arredonda para o par mais próximo (0,25 → 0,2)."""
    from decimal import Decimal, ROUND_HALF_UP
    q = Decimal(1).scaleb(-casas)
    return float(Decimal(repr(float(valor))).quantize(q, rounding=ROUND_HALF_UP))


def arredondar_anvisa(valor: Any, unidade: str = "g") -> float:
    """
    Arredondamento da IN 75/2020, Anexo III:
      - 10 ou mais ............ número inteiro
      - de 1 a menos de 10 .... uma casa decimal
      - menos de 1 (g, kcal) .. uma casa decimal
      - menos de 1 (mg, mcg) .. duas casas decimais
    """
    v = _safe_float(valor)
    if v <= 0:
        return 0.0
    if v >= 10:
        return _half_up(v, 0)
    if v >= 1:
        r = _half_up(v, 1)
        return _half_up(r, 0) if r >= 10 else r
    return _half_up(v, 2 if unidade in ("mg", "mcg") else 1)


def formatar_rotulo(valor: Any, unidade: str) -> str:
    """Texto do valor como vai no rótulo: sem zeros desnecessários, vírgula decimal."""
    if valor is None:
        return "Não informado"
    v = _safe_float(valor)
    if v == int(v):
        txt = str(int(v))
    else:
        txt = f"{v:.2f}".rstrip("0").rstrip(".").replace(".", ",")
    return f"{txt} {unidade}"


def smart_round(valor, contexto: str = "display"):
    """
    contexto='display': retorna string formatada sem zeros desnecessários (até 4 decimais)
    contexto='calc': retorna float arredondado a 4 casas decimais
    """
    v = round(_safe_float(valor), 4)
    if contexto == "calc":
        return v
    if v == int(v):
        return str(int(v))
    return f"{v:.4f}".rstrip("0").rstrip(".")


def convert_to_grams(valor: float, unidade: str, densidade: float = 1.0,
                     total_receita_g: float | None = None) -> float:
    """
    Converte quantidade para gramas.

    unidade: 'g' | 'mL' | '%'
    densidade: g/mL — usado quando unidade='mL'
    total_receita_g: peso total da receita (Modo A para %)
                     None → Modo B: 1% = 1g (base 100g)
    """
    valor = _safe_float(valor)
    if unidade == "g":
        return valor
    if unidade == "mL":
        return round(valor * _safe_float(densidade, 1.0), 4)
    if unidade == "%":
        if total_receita_g is not None and _safe_float(total_receita_g) > 0:
            return round((valor / 100.0) * _safe_float(total_receita_g), 4)
        return valor  # Modo B: 1% = 1g
    return valor


def calcular_composicao_receita(
    ingredientes: list[dict],
    receita_id: int | None = None,
    validar_taco: bool = True,
) -> dict:
    """
    Calcula composição TOTAL somando contribuições de cada ingrediente.

    Args:
        ingredientes: lista de dicts com:
            - composicao_100g: dict com todos os nutrientes
            - quantidade_gramas: float
            - nome: str
            - fonte: str
        receita_id: id da receita para registro de auditoria (opcional)
        validar_taco: se True, executa validação TACO×TBCA em ingredientes TACO

    Returns:
        dict com:
            - nutrientes: {nutriente: valor_total_g_receita}
            - peso_total_gramas: float
            - ingredientes_detalhes: lista com contribuição individual + metadados de fonte
            - alertas_validacao: lista de alertas para exibição na UI
    """
    from modules.database import NUTRIENTES

    # Importação lazy para evitar dependência circular
    _validator = None
    if validar_taco:
        try:
            from modules import validator as _validator
        except ImportError:
            _validator = None

    totais = {n: 0.0 for n in NUTRIENTES}
    peso_total = 0.0
    detalhes = []
    alertas  = []

    for ing in ingredientes:
        comp_original = ing.get("composicao_100g", {})
        qtd   = _safe_float(ing.get("quantidade_gramas", 0))
        nome  = ing.get("nome", "?")
        fonte = ing.get("fonte", "?")

        # ── Validação TACO×TBCA ───────────────────────────────────────────────
        metadados_validacao = {
            "triangulacao_aplicada": False,
            "nivel_confianca":       "—",
            "match_tbca":            "",
            "score_match":           0.0,
            "nutrientes_corrigidos": [],
            "fonte_primaria":        fonte,
        }

        comp = comp_original  # default: usar TACO sem modificação

        if _validator and fonte == "TACO" and comp_original:
            try:
                inconsistencias = _validator.detectar_inconsistencias(nome, comp_original)
                if inconsistencias["inconsistente"]:
                    resultado_tri = _validator.triangular_com_tbca(
                        nome, comp_original, inconsistencias
                    )
                    comp = resultado_tri["dados_finais"]

                    metadados_validacao.update({
                        "triangulacao_aplicada": resultado_tri["triangulacao_aplicada"],
                        "nivel_confianca":       resultado_tri["nivel_confianca"],
                        "match_tbca":            resultado_tri["match_tbca"],
                        "score_match":           resultado_tri["score_match"],
                        "nutrientes_corrigidos": resultado_tri["nutrientes_corrigidos"],
                        "divergencias":          resultado_tri["divergencias"],
                        "fonte_primaria":        "TACO+TBCA" if resultado_tri["triangulacao_aplicada"] else "TACO",
                        "motivo_inconsistencia": inconsistencias["motivo"],
                    })

                    if resultado_tri["triangulacao_aplicada"]:
                        alertas.append({
                            "ingrediente":   nome,
                            "motivo":        inconsistencias["motivo"],
                            "match_tbca":    resultado_tri["match_tbca"],
                            "score_match":   resultado_tri["score_match"],
                            "corrigidos":    resultado_tri["nutrientes_corrigidos"],
                            "nivel":         resultado_tri["nivel_confianca"],
                            "divergencias":  resultado_tri["divergencias"],
                        })

                        if receita_id is not None:
                            _validator.registrar_auditoria(
                                receita_id, nome, resultado_tri
                            )
            except Exception as e:
                print(f"[calculator] Validação falhou para '{nome}': {e}")

        # ── Cálculo da contribuição ───────────────────────────────────────────
        fator  = qtd / 100.0
        contrib = {}
        # Ficha só com kcal (ou só com kJ): completa a outra unidade, senão o
        # total em kJ da receita sairia menor que o equivalente em kcal.
        kcal_ing = _safe_float(comp.get("energia_kcal"))
        kj_ing   = _safe_float(comp.get("energia_kj"))
        if kcal_ing > 0 and kj_ing <= 0:
            comp = {**comp, "energia_kj": kcal_ing * KCAL_PARA_KJ}
        elif kj_ing > 0 and kcal_ing <= 0:
            comp = {**comp, "energia_kcal": kj_ing / KCAL_PARA_KJ}
        for nutriente in NUTRIENTES:
            val          = _safe_float(comp.get(nutriente, 0))
            contribuicao = val * fator
            contrib[nutriente] = contribuicao
            totais[nutriente] += contribuicao

        peso_total += qtd
        detalhes.append({
            "nome":                  nome,
            "fonte":                 fonte,
            "quantidade_gramas":     qtd,
            "composicao_100g":       comp,
            "contribuicao":          contrib,
            # campos extras de unidade original (passados pelo app.py)
            "unidade_original":      ing.get("unidade_original", "g"),
            "quantidade_original":   ing.get("quantidade_original"),
            "densidade_utilizada":   ing.get("densidade_utilizada", 1.0),
            # metadados de rastreabilidade de fonte
            "validacao":             metadados_validacao,
        })

    return {
        "nutrientes":            totais,
        "peso_total_gramas":     peso_total,
        "ingredientes_detalhes": detalhes,
        "alertas_validacao":     alertas,
    }


def calcular_por_100g_produto(composicao_total: dict, peso_total_gramas: float) -> dict:
    """
    Normaliza composição total para base de 100g do produto final.
    """
    if peso_total_gramas <= 0:
        return {n: 0.0 for n in composicao_total}

    fator = 100.0 / peso_total_gramas
    return {n: v * fator for n, v in composicao_total.items()}


def calcular_por_porcao(composicao_total: dict, peso_total_gramas: float, porcao_gramas: float) -> dict:
    """
    Calcula valores para a porção informada.
    """
    if peso_total_gramas <= 0 or porcao_gramas <= 0:
        return {n: 0.0 for n in composicao_total}

    fator = porcao_gramas / peso_total_gramas
    return {n: v * fator for n, v in composicao_total.items()}


def calcular_vd(por_porcao: dict, acucares_adicionados_porcao: float | None = None,
                gordura_saturada_porcao: float | None = None,
                gordura_trans_porcao: float | None = None) -> dict:
    """
    Calcula %VD (inteiro) a partir dos valores declarados na porção.
    Os três parâmetros opcionais existem por compatibilidade; quando omitidos,
    os valores são lidos de por_porcao.
    """
    valores = dict(por_porcao)
    if acucares_adicionados_porcao is not None:
        valores["acucares_adicionados"] = acucares_adicionados_porcao
    if gordura_saturada_porcao is not None:
        valores["gordura_saturada"] = gordura_saturada_porcao
    if gordura_trans_porcao is not None:
        valores["gordura_trans"] = gordura_trans_porcao

    vd: dict = {}
    for nutriente in ("energia_kcal", "carboidrato", "acucares_adicionados",
                      "proteina", "lipideos", "gordura_saturada", "gordura_trans",
                      "fibra_alimentar", "sodio"):
        ref = VALORES_DIARIOS[nutriente]
        val = _safe_float(valores.get(nutriente, 0))
        vd[nutriente] = int(_half_up(val / ref * 100, 0))

    # Açúcares totais: não há Valor Diário estabelecido
    vd["acucares_totais"] = None
    return vd


def aplicar_arredondamentos(valores: dict, porcao_gramas: float = 100.0) -> dict:
    """
    Aplica o arredondamento da IN 75/2020 (Anexo III) às linhas do rótulo.
    Valores None (não informados) são preservados. Os demais nutrientes do
    dicionário (micronutrientes da aba P&D) não são alterados.
    """
    r = dict(valores)
    for nutriente, unidade in UNIDADES_ROTULO.items():
        if nutriente not in r or r[nutriente] is None:
            continue
        r[nutriente] = arredondar_anvisa(r[nutriente], unidade)
    return r


def _zerar_nao_significativos(bruto_100g: dict, bruto_porcao: dict,
                              por_100g: dict, por_porcao: dict) -> None:
    """
    Declara zero quando o valor em 100 g E na porção é não significativo
    (IN 75/2020, Anexo IV), respeitando as condições do próprio anexo:
      - gorduras totais só zeram se saturadas e trans também forem zero;
      - carboidratos só zeram se os açúcares totais também forem zero;
      - açúcares totais só zeram se não houver açúcares adicionados.
    """
    def _insignificante(nutriente: str) -> bool:
        limite = LIMITES_NAO_SIGNIFICATIVOS[nutriente]
        v100, vpor = bruto_100g.get(nutriente), bruto_porcao.get(nutriente)
        if v100 is None or vpor is None:
            return False
        return _safe_float(v100) <= limite and _safe_float(vpor) <= limite

    def _presente(nutriente: str) -> bool:
        """Há valor informado e ele é significativo."""
        return bruto_100g.get(nutriente) is not None and not _insignificante(nutriente)

    tem_adicionado = (_safe_float(bruto_100g.get("acucares_adicionados")) > 0
                      or _safe_float(bruto_porcao.get("acucares_adicionados")) > 0)
    depende_de = {
        "lipideos":        lambda: _presente("gordura_saturada") or _presente("gordura_trans"),
        "carboidrato":     lambda: _presente("acucares_totais") or tem_adicionado,
        "acucares_totais": lambda: tem_adicionado,
    }

    for nutriente in LIMITES_NAO_SIGNIFICATIVOS:
        if not _insignificante(nutriente):
            continue
        if nutriente in depende_de and depende_de[nutriente]():
            continue
        por_100g[nutriente] = 0.0
        por_porcao[nutriente] = 0.0
        if nutriente == "energia_kcal":
            por_100g["energia_kj"] = 0.0
            por_porcao["energia_kj"] = 0.0


def montar_tabela_rotulo(
    nome_produto:    str,
    porcao_gramas:   float,
    peso_total:      float,
    nutrientes_totais: dict,
    acucares_adicionados_total: float = 0.0,
    gordura_saturada_total: float = 0.0,
    gordura_trans_total: float = 0.0,
    num_porcoes: int | None = None,
    medida_caseira: str | None = None,
    acucares_totais_total: float | None = None,
) -> dict:
    """
    Função principal: monta a tabela completa para o rótulo ANVISA.

    acucares_totais_total: açúcares totais da receita inteira, em gramas.
        A TACO não traz esse dado; quando não informado (None) a linha sai como
        "Não informado" em vez de um valor estimado.

    Returns:
        dict com:
            - por_100g: dict (valores arredondados)
            - por_porcao: dict (valores arredondados)
            - vd: dict (%VD por nutriente)
            - info: dict (metadados para o rótulo)
    """
    por_100g_raw   = calcular_por_100g_produto(nutrientes_totais, peso_total)
    por_porcao_raw = calcular_por_porcao(nutrientes_totais, peso_total, porcao_gramas)

    f100 = 100.0 / peso_total if peso_total > 0 else 0.0
    fpor = porcao_gramas / peso_total if peso_total > 0 else 0.0

    informados = {
        "gordura_saturada":     _safe_float(gordura_saturada_total),
        "gordura_trans":        _safe_float(gordura_trans_total),
        "acucares_adicionados": _safe_float(acucares_adicionados_total),
    }
    for chave, total in informados.items():
        por_100g_raw[chave]   = total * f100
        por_porcao_raw[chave] = total * fpor

    if acucares_totais_total is None:
        por_100g_raw["acucares_totais"]   = None
        por_porcao_raw["acucares_totais"] = None
    else:
        por_100g_raw["acucares_totais"]   = _safe_float(acucares_totais_total) * f100
        por_porcao_raw["acucares_totais"] = _safe_float(acucares_totais_total) * fpor

    por_100g   = aplicar_arredondamentos(por_100g_raw,   100)
    por_porcao = aplicar_arredondamentos(por_porcao_raw, porcao_gramas)
    _zerar_nao_significativos(por_100g_raw, por_porcao_raw, por_100g, por_porcao)

    # Texto da linha de gorduras trans (mantido para compatibilidade)
    por_porcao["gordura_trans_rotulo"] = formatar_rotulo(por_porcao["gordura_trans"], "g")

    vd = calcular_vd(por_porcao)

    return {
        "por_100g":     por_100g,
        "por_porcao":   por_porcao,
        "por_100g_bruto": por_100g_raw,
        "vd":           vd,
        "info": {
            "nome_produto":   nome_produto,
            "porcao_gramas":  porcao_gramas,
            "peso_total":     peso_total,
            "num_porcoes":    num_porcoes,
            "medida_caseira": medida_caseira or "",
        },
    }


def calcular_contribuicoes_percentuais(detalhes: list[dict]) -> list[dict]:
    """
    Calcula % de contribuição de cada ingrediente no total de cada nutriente.
    """
    from modules.database import NUTRIENTES

    # Calcular totais por nutriente
    totais_nut = {n: 0.0 for n in NUTRIENTES}
    for ing in detalhes:
        for n in NUTRIENTES:
            totais_nut[n] += _safe_float(ing["contribuicao"].get(n, 0))

    # Calcular percentuais
    resultado = []
    for ing in detalhes:
        contrib_perc = {}
        for n in NUTRIENTES:
            total = totais_nut[n]
            val   = _safe_float(ing["contribuicao"].get(n, 0))
            contrib_perc[n] = round((val / total) * 100, 1) if total > 0 else 0.0

        resultado.append({
            **ing,
            "contribuicao_percentual": contrib_perc,
        })

    return resultado


# ─── LIMITAÇÕES CONHECIDAS ─────────────────────────────────────────────────────
# 1. Gordura saturada e trans não são diferenciadas na TACO — devem ser informadas
#    manualmente pelo usuário ou calculadas por formulação específica.
# 2. Açúcares adicionados ≠ açúcares totais — a TACO não diferencia; usuário deve
#    informar separadamente a quantidade de açúcar adicionado na formulação.
# 3. Açúcares totais não constam na TACO — devem ser informados pelo usuário;
#    sem esse dado a linha do rótulo sai como "Não informado" (nunca estimada).
# 4. Perdas nutricionais por processamento térmico (cozimento, pasteurização, etc.)
#    NÃO são calculadas — a TACO fornece valores de alimentos crus ou em estado padrão.
# 5. Medida caseira deve ser informada manualmente — não é calculada automaticamente.
# 6. Variações sazonais e de origem dos ingredientes não são contempladas pela TACO.
# 7. OBRIGATÓRIO: validação por nutricionista habilitado (CRN) antes de uso em
#    rótulo comercial para venda. Este sistema é uma ferramenta de P&D, não substitui
#    análise laboratorial nem laudo técnico obrigatório pela legislação.
