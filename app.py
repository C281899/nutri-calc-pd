"""
app.py — NutriCalc P&D
Aplicação principal Streamlit. Interface unificada para cálculo de tabelas
nutricionais no padrão ANVISA (RDC 429/2020 e IN 75/2020).

Execução:
    streamlit run app.py
"""

import os
import re
import sys
import sqlite3
from datetime import datetime
from html import escape

import streamlit as st

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

TACO_DB = os.path.join(ROOT_DIR, "taco.db")
APP_DB  = os.path.join(ROOT_DIR, "nutri_calc.db")
TBCA_DB = os.path.join(ROOT_DIR, "tbca.db")


# ══════════════════════════════════════════════════════════════════════════════
#  INICIALIZAÇÃO DOS BANCOS
# ══════════════════════════════════════════════════════════════════════════════

def _carregar_setup(nome_arquivo: str):
    """Importa dinamicamente um script de data/ (taco_setup.py, tbca_setup.py)."""
    import importlib.util
    caminho = os.path.join(ROOT_DIR, "data", nome_arquivo)
    spec = importlib.util.spec_from_file_location(nome_arquivo[:-3], caminho)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _tabela_tem_dados(caminho_db: str, tabela: str) -> bool:
    """True se o banco existe, abre e a tabela tem ao menos uma linha."""
    if not os.path.exists(caminho_db):
        return False
    try:
        con = sqlite3.connect(caminho_db)
        try:
            return con.execute(f"SELECT COUNT(*) FROM {tabela}").fetchone()[0] > 0
        finally:
            con.close()
    except sqlite3.Error:
        return False


def _tabela_existe(caminho_db: str, tabela: str) -> bool:
    if not os.path.exists(caminho_db):
        return False
    try:
        con = sqlite3.connect(caminho_db)
        try:
            return con.execute(
                "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (tabela,)
            ).fetchone() is not None
        finally:
            con.close()
    except sqlite3.Error:
        return False


@st.cache_resource(show_spinner="Preparando NutriCalc P&D (primeira execução)...")
def preparar_bancos() -> dict:
    """
    Cria/valida os bancos UMA vez por processo do servidor.

    - cache_resource serializa a chamada: duas pessoas abrindo o app ao mesmo
      tempo depois de um reinício não disparam duas criações concorrentes;
    - confere o conteúdo, não só a existência do arquivo — um banco deixado
      pela metade por uma inicialização interrompida é refeito.
    """
    avisos = []

    if not _tabela_tem_dados(TACO_DB, "alimentos_taco"):
        _carregar_setup("taco_setup.py").criar_taco_db()
    else:
        # banco de uma versão anterior: inclui itens novos do suplemento
        try:
            _carregar_setup("taco_setup.py").completar_suplemento()
        except Exception as e:
            avisos.append(f"Não foi possível atualizar a base de alimentos ({e}).")
    if not _tabela_existe(APP_DB, "receitas"):
        _carregar_setup("taco_setup.py").criar_app_db()
    if not _tabela_tem_dados(TBCA_DB, "tbca_alimentos"):
        try:
            _carregar_setup("tbca_setup.py").criar_tbca_db()
        except Exception as e:
            avisos.append(f"TBCA não carregada ({e}). Validação cruzada desativada.")

    from modules import database as _db
    _db.migrar_banco()

    con = sqlite3.connect(TACO_DB)
    try:
        total = con.execute("SELECT COUNT(*) FROM alimentos_taco").fetchone()[0]
    finally:
        con.close()
    return {"total_alimentos": total, "avisos": avisos}


st.set_page_config(
    page_title="NutriCalc P&D",
    page_icon="🔬",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        "About": "**NutriCalc P&D** v2.2 · Cálculo de Tabelas Nutricionais · ANVISA RDC 429/2020",
    },
)

from modules import style as sty
sty.inject_global_css()

try:
    _estado_bancos = preparar_bancos()
except Exception as e:
    st.error(f"Erro ao inicializar o banco de dados: {e}")
    st.stop()

import pandas as pd

from modules import database as db
from modules import calculator as calc
from modules import parser as prs
from modules import excel_generator as xls
from modules.database import NUTRIENTES

_db_count = _estado_bancos["total_alimentos"]

sty.render_page_header(_db_count)
for _aviso in _estado_bancos["avisos"]:
    st.warning(_aviso)

# Banner de aviso sobre armazenamento efêmero
st.markdown("""
<div style="
    background:#1A1200;
    border:0.5px solid #5A3A00;
    border-radius:8px;
    padding:10px 14px;
    margin-bottom:16px;
    font-size:11px;
    color:#EF9F27;
    line-height:1.6;
">
  <strong style="color:#FAC775;">Armazenamento temporário:</strong>
  receitas e ingredientes salvos nesta sessão serão perdidos quando o servidor reiniciar
  (após ~1h sem uso no Streamlit Cloud).
  <strong>Exporte sempre seu trabalho em Excel</strong> ou use o botão
  <em>Baixar backup</em> na aba "Receitas Salvas".
</div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
#  ESTADO DA SESSÃO
# ══════════════════════════════════════════════════════════════════════════════

# Campos do formulário "Dados do Produto" (chave do widget → valor inicial)
FORM_PADRAO = {
    "form_nome":       "",
    "form_porcao":     30.0,
    "form_nporcoes":   1,
    "form_medida":     "",
    "form_gord_sat":   0.0,
    "form_gord_trans": 0.0,
    "form_acucar_ad":  0.0,
    "form_acucar_tot": None,   # None = não informado
}

SEM_SELECAO = "-- Selecionar --"


def _init_session():
    defaults = {
        "ingredientes":          [],
        "resultado_calculo":     None,
        "receita_editando_id":   None,
        "upload_preview":        None,
        "preview_ver":           0,     # muda a cada novo arquivo → widgets novos
        "forn_editando_id":      None,
        "forn_dados_extraidos":  {},
        "forn_form_ver":         0,     # muda quando o formulário deve ser recarregado
        "ficha_unico":           None,
        "form_total_receita_g":  1000.0,
        "_uid_seq":              0,
        "_flash":                [],
        **FORM_PADRAO,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


_init_session()


def _flash(mensagem: str, tipo: str = "success"):
    """Guarda um aviso para ser mostrado na próxima execução do script.
    (st.success seguido de st.rerun() nunca chega a aparecer na tela.)"""
    st.session_state["_flash"].append((tipo, mensagem))


def _mostrar_flash():
    pendentes = st.session_state.get("_flash") or []
    for tipo, msg in pendentes:
        getattr(st, tipo, st.info)(msg)
    st.session_state["_flash"] = []


def _novo_ingrediente(nome, fonte, fonte_id, qtd_gramas, comp,
                      unidade="g", qtd_original=None, densidade=1.0) -> dict:
    """Monta um ingrediente da receita com identificador próprio (uid).

    O uid dá a cada linha da lista um widget que é só dela. Antes os campos de
    quantidade eram identificados pela POSIÇÃO na lista: ao remover um item, os
    de baixo herdavam o valor digitado no item de cima.
    """
    st.session_state["_uid_seq"] += 1
    return {
        "uid":                 st.session_state["_uid_seq"],
        "nome":                nome,
        "fonte":               fonte,
        "fonte_id":            fonte_id,
        "quantidade_gramas":   float(qtd_gramas),
        "composicao_100g":     comp,
        "unidade_original":    unidade,
        "quantidade_original": float(qtd_gramas if qtd_original is None else qtd_original),
        "densidade_utilizada": densidade,
    }


def _limpar_receita(manter_cabecalho: bool = False):
    st.session_state.ingredientes        = []
    st.session_state.resultado_calculo   = None
    st.session_state.receita_editando_id = None
    st.session_state.upload_preview      = None
    st.session_state.ficha_unico         = None
    if not manter_cabecalho:
        for chave, valor in FORM_PADRAO.items():
            st.session_state[chave] = valor


# ── Callbacks (rodam ANTES do script, então podem alterar campos do formulário) ──

def _cb_nova_receita():
    _limpar_receita()


def _cb_limpar_ingredientes():
    _limpar_receita(manter_cabecalho=True)


def _cb_remover_ingrediente(uid: int):
    st.session_state.ingredientes = [
        i for i in st.session_state.ingredientes if i["uid"] != uid
    ]


def _cb_abrir_receita(mapa_rec: dict):
    escolha = st.session_state.get("sel_abrir")
    rid = mapa_rec.get(escolha)
    receita = db.get_receita_completa(rid) if rid else None
    if not receita:
        _flash("Receita não encontrada.", "error")
        return

    carregados, perdidos, religados = [], [], []
    for ing in receita["ingredientes"]:
        achado = db.resolver_ingrediente(ing["fonte"], ing["fonte_id"], ing["nome_ingrediente"])
        if not achado:
            perdidos.append(ing["nome_ingrediente"])
            continue
        if achado["fonte"] != ing["fonte"]:
            religados.append(f"{ing['nome_ingrediente']} ({achado['fonte']})")
        carregados.append(_novo_ingrediente(
            nome=achado["nome"],
            fonte=achado["fonte"],
            fonte_id=achado["fonte_id"],
            qtd_gramas=ing["quantidade_gramas"],
            comp=achado["composicao_100g"],
            unidade=ing.get("unidade_original") or "g",
            qtd_original=ing.get("quantidade_original"),
            densidade=ing.get("densidade_utilizada") or 1.0,
        ))

    st.session_state.ingredientes        = carregados
    st.session_state.resultado_calculo   = None
    st.session_state.receita_editando_id = rid
    st.session_state.upload_preview      = None
    st.session_state.ficha_unico         = None
    st.session_state.form_nome       = receita["nome_produto"] or ""
    st.session_state.form_porcao     = float(receita["porcao_gramas"] or 30.0)
    st.session_state.form_nporcoes   = int(receita["num_porcoes"] or 1)
    st.session_state.form_medida     = receita.get("medida_caseira") or ""
    st.session_state.form_gord_sat   = float(receita.get("gordura_saturada_total") or 0.0)
    st.session_state.form_gord_trans = float(receita.get("gordura_trans_total") or 0.0)
    st.session_state.form_acucar_ad  = float(receita.get("acucares_adicionados_total") or 0.0)
    st.session_state.form_acucar_tot = receita.get("acucares_totais_total")
    st.session_state.sel_abrir       = SEM_SELECAO

    _flash(f"Receita '{receita['nome_produto']}' carregada — abra a aba 'Nova Receita'.")
    if perdidos:
        _flash(
            "Ingredientes que não existem mais no banco e ficaram FORA da receita: "
            + ", ".join(perdidos)
            + ". Cadastre-os de novo em 'Ingredientes de Fornecedor' ou restaure o backup.",
            "error",
        )
    if religados:
        _flash("Ingredientes localizados em outra fonte: " + ", ".join(religados), "warning")


def _cb_editar_fornecedor(mapa_forn: dict):
    fid = mapa_forn.get(st.session_state.get("sel_forn_editar"))
    if fid:
        st.session_state.forn_editando_id = fid
        st.session_state.forn_form_ver   += 1
        st.session_state.sel_forn_editar  = SEM_SELECAO


def _cb_cancelar_edicao_fornecedor():
    st.session_state.forn_editando_id     = None
    st.session_state.forn_dados_extraidos = {}
    st.session_state.forn_form_ver       += 1


# ══════════════════════════════════════════════════════════════════════════════
#  CÁLCULO
# ══════════════════════════════════════════════════════════════════════════════

def _assinatura_entrada(cabecalho: dict) -> tuple:
    """Resumo de tudo que influencia o resultado. Se mudar, o resultado em tela
    está desatualizado e precisa ser refeito."""
    return (
        tuple(sorted((k, repr(v)) for k, v in cabecalho.items())),
        tuple(
            (i["uid"], round(float(i["quantidade_gramas"]), 6), i["fonte"], i["fonte_id"])
            for i in st.session_state.ingredientes
        ),
    )


def _validar_entrada(cabecalho: dict) -> list[str]:
    erros = []
    if not cabecalho["nome"]:
        erros.append("Informe o nome do produto.")
    if not cabecalho["porcao_g"] or cabecalho["porcao_g"] <= 0:
        erros.append("Informe o tamanho da porção em gramas.")
    if not st.session_state.ingredientes:
        erros.append("Adicione pelo menos um ingrediente antes de calcular.")
    elif sum(i["quantidade_gramas"] for i in st.session_state.ingredientes) <= 0:
        erros.append("A soma das quantidades dos ingredientes precisa ser maior que zero.")
    return erros


def _calcular(cabecalho: dict, registrar_auditoria: bool) -> dict:
    """Executa o cálculo completo e devolve o dicionário de resultado."""
    resultado_bruto = calc.calcular_composicao_receita(
        st.session_state.ingredientes,
        # auditoria só no clique do botão — o recálculo automático não deve
        # encher a tabela de auditoria com linhas repetidas
        receita_id=st.session_state.get("receita_editando_id") if registrar_auditoria else None,
        validar_taco=True,
    )
    peso_total = resultado_bruto["peso_total_gramas"]
    tabela = calc.montar_tabela_rotulo(
        nome_produto=cabecalho["nome"],
        porcao_gramas=cabecalho["porcao_g"],
        peso_total=peso_total,
        nutrientes_totais=resultado_bruto["nutrientes"],
        acucares_adicionados_total=cabecalho["acucar_ad"],
        gordura_saturada_total=cabecalho["gord_sat"],
        gordura_trans_total=cabecalho["gord_trans"],
        num_porcoes=cabecalho["num_porcoes"],
        medida_caseira=cabecalho["medida"] or None,
        acucares_totais_total=cabecalho["acucar_tot"],
    )

    # Conferências de coerência dos campos informados à mão
    avisos = []
    tot = resultado_bruto["nutrientes"]
    if cabecalho["gord_sat"] + cabecalho["gord_trans"] > tot.get("lipideos", 0) + 0.05:
        avisos.append(
            f"Gordura saturada + trans informadas ({cabecalho['gord_sat'] + cabecalho['gord_trans']:.2f} g) "
            f"superam as gorduras totais calculadas da receita ({tot.get('lipideos', 0):.2f} g)."
        )
    if cabecalho["acucar_tot"] is not None:
        if cabecalho["acucar_ad"] > cabecalho["acucar_tot"] + 0.05:
            avisos.append("Açúcares adicionados informados superam os açúcares totais informados.")
        if cabecalho["acucar_tot"] > tot.get("carboidrato", 0) + 0.05:
            avisos.append(
                f"Açúcares totais informados ({cabecalho['acucar_tot']:.2f} g) superam os "
                f"carboidratos calculados da receita ({tot.get('carboidrato', 0):.2f} g)."
            )
    elif cabecalho["acucar_ad"] > tot.get("carboidrato", 0) + 0.05:
        avisos.append("Açúcares adicionados informados superam os carboidratos calculados da receita.")
    if cabecalho["porcao_g"] > peso_total + 1e-6:
        avisos.append(
            f"A porção ({cabecalho['porcao_g']:g} g) é maior que o peso total da receita "
            f"({peso_total:g} g)."
        )

    return {
        "tabela":            tabela,
        "detalhes":          calc.calcular_contribuicoes_percentuais(
                                 resultado_bruto["ingredientes_detalhes"]),
        "peso_total":        peso_total,
        "nome_produto":      cabecalho["nome"],
        "porcao_g":          cabecalho["porcao_g"],
        "num_porcoes":       cabecalho["num_porcoes"],
        "medida_caseira":    cabecalho["medida"] or None,
        "alertas_validacao": resultado_bruto.get("alertas_validacao", []),
        "avisos":            avisos,
        "extras": {
            "gordura_saturada_total":     cabecalho["gord_sat"],
            "gordura_trans_total":        cabecalho["gord_trans"],
            "acucares_adicionados_total": cabecalho["acucar_ad"],
            "acucares_totais_total":      cabecalho["acucar_tot"],
        },
        "assinatura":        _assinatura_entrada(cabecalho),
    }


def _nome_arquivo_seguro(nome: str) -> str:
    """Nome de arquivo sem caracteres que o Windows/Excel rejeitam."""
    limpo = re.sub(r'[\\/:*?"<>|\r\n\t]+', " ", nome).strip()
    limpo = re.sub(r"\s+", "_", limpo)[:80]
    return limpo or "receita"


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    sty.render_sidebar(_db_count)

_mostrar_flash()

# ── Abas principais ────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["Nova Receita", "Ingredientes de Fornecedor", "Receitas Salvas"])


# ══════════════════════════════════════════════════════════════════════════════
#  ABA 1 — NOVA RECEITA
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    col_esq, col_dir = st.columns([3, 1])

    with col_esq:
        sty.render_section_label("Dados do Produto")

    with col_dir:
        if st.session_state.receita_editando_id:
            st.info(f"Editando receita #{st.session_state.receita_editando_id}")
            st.button("Nova Receita", width="stretch", on_click=_cb_nova_receita,
                      key="btn_nova_receita")

    c1, c2, c3 = st.columns([3, 1.5, 1.5])
    nome_produto = c1.text_input(
        "Nome do Produto *",
        placeholder="Ex.: Bolo de Chocolate Diet",
        key="form_nome",
    )
    porcao_g = c2.number_input(
        "Tamanho da Porção (g) *",
        min_value=0.1, max_value=10000.0,
        step=1.0, format="%.1f",
        key="form_porcao",
    )
    num_porcoes = c3.number_input(
        "Nº de Porções por Embalagem",
        min_value=1, max_value=9999, step=1,
        key="form_nporcoes",
    )
    medida_caseira = st.text_input(
        "Medida Caseira (opcional)",
        placeholder="Ex.: 1 fatia, 2 colheres de sopa, 1 unidade",
        key="form_medida",
    )

    with st.expander("Campos Adicionais ANVISA — Gorduras Saturada e Trans, Açúcares Totais e Adicionados"):
        st.caption(
            "Estes valores não constam na TACO. Informe o total presente NA RECEITA INTEIRA (g), "
            "com base na formulação. Açúcares totais em branco saem como 'Não informado' no rótulo."
        )
        c4, c5, c6, c7 = st.columns(4)
        gordura_saturada_total = c4.number_input(
            "Gordura Saturada na Receita (g)",
            min_value=0.0, max_value=100000.0, step=0.1,
            help="Soma das gorduras saturadas de todos os ingredientes",
            key="form_gord_sat",
        )
        gordura_trans_total = c5.number_input(
            "Gordura Trans na Receita (g)",
            min_value=0.0, max_value=100000.0, step=0.01,
            key="form_gord_trans",
        )
        acucares_totais_total = c6.number_input(
            "Açúcares Totais na Receita (g)",
            min_value=0.0, max_value=100000.0, step=0.1,
            value=None, placeholder="Não informado",
            help="Açúcares naturalmente presentes + adicionados. A TACO não traz este dado.",
            key="form_acucar_tot",
        )
        acucares_adicionados_total = c7.number_input(
            "Açúcares Adicionados na Receita (g)",
            min_value=0.0, max_value=100000.0, step=0.1,
            help="Quantidade de açúcar de adição (sacarose, mel, xarope, etc.)",
            key="form_acucar_ad",
        )

    cabecalho = {
        "nome":        (nome_produto or "").strip(),
        "porcao_g":    float(porcao_g or 0),
        "num_porcoes": int(num_porcoes or 1),
        "medida":      (medida_caseira or "").strip(),
        "gord_sat":    float(gordura_saturada_total or 0),
        "gord_trans":  float(gordura_trans_total or 0),
        "acucar_ad":   float(acucares_adicionados_total or 0),
        "acucar_tot":  None if acucares_totais_total is None else float(acucares_totais_total),
    }

    sty.render_divider()
    sty.render_section_label("Ingredientes")

    modo = st.radio(
        "Modo de entrada:",
        ["Manual (item a item)", "Upload de Ficha Técnica"],
        horizontal=True,
        label_visibility="collapsed",
    )

    # ── Modo A — Entrada Manual ───────────────────────────────────────────────
    if modo == "Manual (item a item)":
        st.markdown("**Adicionar Ingrediente**")
        c_busca, c_qtd, c_unit, c_btn = st.columns([4, 1.5, 1, 2])

        with c_busca:
            termo = st.text_input(
                "Buscar ingrediente",
                placeholder="Digite o nome do ingrediente e tecle Enter...",
                key="busca_ingrediente",
                label_visibility="collapsed",
            )

        opcoes_busca = []
        mapa_busca   = {}
        if termo and len(termo.strip()) >= 3:
            candidatos = db.buscar_ingrediente(termo, limite=15)
            for c in candidatos:
                label = f"[{c['fonte']}] {c['nome']} ({c['score']:.0f}%)"
                if label in mapa_busca:          # nomes repetidos entre fontes
                    label += f" · id {c['id']}"
                opcoes_busca.append(label)
                mapa_busca[label] = c

        ingrediente_selecionado = None
        if opcoes_busca:
            escolha = st.selectbox(
                "Selecionar:",
                opcoes_busca,
                key="sel_ingrediente",
                label_visibility="collapsed",
            )
            ingrediente_selecionado = mapa_busca.get(escolha)
        elif termo and len(termo.strip()) >= 3:
            st.caption("Nenhum ingrediente encontrado. Verifique o nome ou cadastre em 'Ingredientes de Fornecedor'.")
        elif termo and len(termo.strip()) < 3:
            st.caption("Digite pelo menos 3 caracteres para buscar.")

        with c_unit:
            unidade = st.selectbox(
                "Unidade",
                ["g", "mL", "%"],
                key="sel_unidade",
                label_visibility="collapsed",
            )

        with c_qtd:
            qtd_manual = st.number_input(
                f"Quantidade ({unidade})",
                min_value=0.0, max_value=100000.0,
                value=100.0, step=1.0,
                format="%.4f",
                key="qtd_manual",
                label_visibility="collapsed",
            )

        with c_btn:
            adicionar_click = st.button("Adicionar", width="stretch", key="btn_adicionar")

        # ── Configuração de unidade selecionada ───────────────────────────────
        densidade_usada = 1.0
        total_receita_g_pct = None

        if unidade == "mL":
            nome_busca_dens = ingrediente_selecionado["nome"] if ingrediente_selecionado else (termo or "")
            dens_achada = db.buscar_densidade_detalhe(nome_busca_dens) if nome_busca_dens else None
            densidade_auto = dens_achada[1] if dens_achada else None
            c_dens1, c_dens2 = st.columns([3, 1])
            with c_dens1:
                if dens_achada:
                    st.caption(
                        f"Densidade da tabela ('{dens_achada[0]}'): **{densidade_auto} g/mL** "
                        "— confira e ajuste ao lado se o seu ingrediente for diferente."
                    )
                else:
                    st.caption("Densidade não encontrada na tabela — informe ao lado (g/mL).")
            with c_dens2:
                # a chave inclui o ingrediente: ao trocar de ingrediente o campo
                # recebe a densidade dele, em vez de manter a do anterior
                densidade_usada = st.number_input(
                    "Densidade (g/mL)",
                    min_value=0.001, max_value=5.0,
                    value=float(densidade_auto or 1.0),
                    step=0.001, format="%.3f",
                    key=f"densidade_manual_{nome_busca_dens}",
                    label_visibility="collapsed",
                )

        elif unidade == "%":
            c_pct1, c_pct2 = st.columns([2, 2])
            with c_pct1:
                modo_pct = st.radio(
                    "Modo %:",
                    ["A — Peso total fixo", "B — Base 100g"],
                    key="radio_modo_pct",
                    horizontal=True,
                    label_visibility="collapsed",
                )
            with c_pct2:
                if "A" in modo_pct:
                    total_receita_g_pct = st.number_input(
                        "Peso total da receita (g)",
                        min_value=1.0, max_value=100000.0,
                        step=1.0, format="%.1f",
                        key="form_total_receita_g",
                        label_visibility="collapsed",
                    )
                else:
                    st.caption("Modo B: 1% = 1g (formulação base 100g)")

            # Indicador de soma % em tempo real
            soma_pct = sum(
                i.get("quantidade_original") or 0
                for i in st.session_state.ingredientes
                if i.get("unidade_original") == "%"
            )
            soma_com_atual = soma_pct + qtd_manual
            if abs(soma_com_atual - 100.0) < 0.01:
                st.success(f"Soma percentual com este item: {soma_com_atual:.4f}% — formulação completa (100%)")
            elif soma_com_atual > 100.0:
                st.error(f"Soma percentual com este item: {soma_com_atual:.4f}% — excede 100%")
            else:
                st.warning(f"Soma percentual com este item: {soma_com_atual:.4f}% — faltam {100 - soma_com_atual:.4f}%")

        if adicionar_click:
            if not ingrediente_selecionado:
                st.warning("Busque e selecione um ingrediente da lista antes de adicionar.")
            elif qtd_manual <= 0:
                st.warning("Informe uma quantidade maior que zero.")
            else:
                comp = db.get_composicao_por_100g(
                    ingrediente_selecionado["fonte"],
                    ingrediente_selecionado["id"],
                )
                if comp:
                    qtd_gramas = calc.convert_to_grams(
                        qtd_manual, unidade, densidade_usada, total_receita_g_pct
                    )
                    st.session_state.ingredientes.append(_novo_ingrediente(
                        nome=ingrediente_selecionado["nome"],
                        fonte=ingrediente_selecionado["fonte"],
                        fonte_id=ingrediente_selecionado["id"],
                        qtd_gramas=qtd_gramas,
                        comp=comp,
                        unidade=unidade,
                        qtd_original=qtd_manual,
                        densidade=densidade_usada if unidade == "mL" else 1.0,
                    ))
                    st.rerun()
                else:
                    st.error("Não foi possível obter a composição deste ingrediente.")

    # ── Modo B — Upload de Ficha Técnica ──────────────────────────────────────
    else:
        arquivo = st.file_uploader(
            "Carregar ficha técnica",
            type=["pdf", "xlsx", "xls", "csv", "txt"],
            label_visibility="collapsed",
            key="upload_ficha",
        )
        _backend = prs.PDF_BACKEND
        _backend_label = "pymupdf (OCR ativo)" if _backend == "pymupdf" else "pdfplumber (sem OCR)"
        st.caption(f"Leitor de PDF ativo: **{_backend_label}**")

        if arquivo is not None:
            if st.button("Extrair Ingredientes", key="btn_extrair"):
                with st.spinner("Analisando arquivo..."):
                    try:
                        ext = arquivo.name.rsplit(".", 1)[-1]
                        extraidos = prs.parse_ficha_tecnica(arquivo.getvalue(), ext)

                        item_unico = next((i for i in extraidos if "composicao_direta" in i), None)
                        if item_unico:
                            if not item_unico["nome"]:
                                item_unico["nome"] = (
                                    arquivo.name.rsplit(".", 1)[0]
                                    .replace("_", " ").replace("-", " ").title()
                                )
                            st.session_state["ficha_unico"] = item_unico
                            st.session_state.upload_preview  = None
                            st.session_state.preview_ver    += 1
                            n_nut = sum(
                                1 for v in item_unico["composicao_direta"].values()
                                if v is not None
                            )
                            st.success(
                                f"Ficha de ingrediente único detectada — "
                                f"{n_nut} nutrientes extraídos. Confirme abaixo."
                            )
                        else:
                            preview = []
                            for item in extraidos:
                                candidatos = db.buscar_ingrediente(item["nome"], limite=5)
                                preview.append({
                                    "nome_original":     item["nome"],
                                    "quantidade_gramas": item["quantidade_gramas"],
                                    "_candidatos":       candidatos,
                                })
                            st.session_state.upload_preview  = preview
                            st.session_state.preview_ver    += 1
                            st.session_state["ficha_unico"]  = None
                            st.success(f"{len(extraidos)} ingredientes encontrados no arquivo.")
                    except ValueError as e:
                        st.error(str(e))
                    except Exception as e:
                        st.error(
                            f"Erro inesperado ao processar arquivo: {type(e).__name__}. "
                            "Tente outro arquivo ou use a entrada manual."
                        )

        # ── Ficha de ingrediente único: confirmar e cadastrar ─────────────────
        if st.session_state.get("ficha_unico"):
            fu   = st.session_state["ficha_unico"]
            comp = fu["composicao_direta"]
            ver  = st.session_state.preview_ver

            sty.render_divider()
            sty.render_section_label("Ficha de Ingrediente Único Detectada")
            st.caption("Verifique e edite os valores (por 100 g) abaixo, depois clique em Cadastrar.")

            LABELS_FU = {
                "energia_kcal":    "Energia (kcal)",
                "energia_kj":      "Energia (kJ)",
                "proteina":        "Proteínas (g)",
                "lipideos":        "Gorduras Totais (g)",
                "carboidrato":     "Carboidratos (g)",
                "fibra_alimentar": "Fibra Alimentar (g)",
                "sodio":           "Sódio (mg)",
                "umidade":         "Umidade (%)",
                "colesterol":      "Colesterol (mg)",
                "calcio":          "Cálcio (mg)",
                "ferro":           "Ferro (mg)",
                "potassio":        "Potássio (mg)",
                "vitamina_c":      "Vitamina C (mg)",
                "vitamina_d":      "Vitamina D (mcg)",
                "vitamina_e":      "Vitamina E (mg)",
                "vitamina_b12":    "Vitamina B12 (mcg)",
            }

            with st.form(f"form_ficha_unico_{ver}"):
                fu_nome = st.text_input("Nome do ingrediente *", value=fu["nome"])
                fu_fab  = st.text_input("Fabricante (opcional)", value="")

                cols_fu = st.columns(4)
                vals_fu: dict = {}
                for i, nut in enumerate(LABELS_FU):
                    v = comp.get(nut)
                    vals_fu[nut] = cols_fu[i % 4].number_input(
                        LABELS_FU[nut],
                        min_value=0.0,
                        value=float(v) if v is not None else 0.0,
                        step=0.01,
                        key=f"fu_{nut}_{ver}",
                    )

                confirmar = st.form_submit_button(
                    "Cadastrar como Ingrediente de Fornecedor",
                    type="primary",
                    width="stretch",
                )

            if confirmar:
                if not fu_nome.strip():
                    st.error("Informe o nome do ingrediente.")
                else:
                    try:
                        # nutrientes que a ficha trouxe mas não aparecem no formulário
                        # (magnésio, zinco...) também são gravados
                        dados_salvar = {
                            **{k: v for k, v in comp.items() if k in NUTRIENTES},
                            "nome_comercial": fu_nome.strip(),
                            "fabricante":     fu_fab.strip() or None,
                            **vals_fu,
                        }
                        iid = db.salvar_ingrediente_fornecedor(dados_salvar)
                        st.session_state["ficha_unico"] = None
                        comp_100g = db.get_composicao_por_100g("FORNECEDOR", iid)
                        if comp_100g:
                            st.session_state.ingredientes.append(_novo_ingrediente(
                                nome=fu_nome.strip(), fonte="FORNECEDOR", fonte_id=iid,
                                qtd_gramas=100.0, comp=comp_100g,
                            ))
                        _flash(
                            f"Ingrediente '{fu_nome.strip()}' cadastrado a partir da ficha técnica "
                            "e incluído na receita com 100 g — ajuste a quantidade na lista."
                        )
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erro ao salvar: {e}")

        # ── Preview editável (multi-ingredientes) ─────────────────────────────
        if st.session_state.upload_preview:
            st.markdown("**Revisar e confirmar ingredientes extraídos:**")
            preview = st.session_state.upload_preview
            ver     = st.session_state.preview_ver   # chaves novas a cada arquivo
            IGNORAR = "-- Ignorar este item --"

            for i, item in enumerate(preview):
                c1, c2, c3 = st.columns([3, 2, 3])
                c1.text_input(
                    "Nome original",
                    value=item["nome_original"],
                    disabled=True,
                    key=f"prev_nome_{ver}_{i}",
                )
                c2.number_input(
                    "Qtd (g)",
                    value=float(item["quantidade_gramas"]),
                    min_value=0.0,
                    step=1.0,
                    format="%.4f",
                    key=f"prev_qtd_{ver}_{i}",
                )
                cands  = item["_candidatos"]
                opcoes = [IGNORAR] + [
                    f"[{c['fonte']}] {c['nome']} ({c['score']:.0f}%) · id {c['id']}" for c in cands
                ]
                c3.selectbox(
                    "Match TACO/Fornecedor",
                    opcoes,
                    index=1 if cands and cands[0]["score"] >= 80 else 0,
                    key=f"prev_match_{ver}_{i}",
                )

            if st.button("Importar Ingredientes Selecionados", key="btn_importar"):
                importados, ignorados = 0, []
                for i, item in enumerate(preview):
                    cands  = item["_candidatos"]
                    opcoes = [IGNORAR] + [
                        f"[{c['fonte']}] {c['nome']} ({c['score']:.0f}%) · id {c['id']}" for c in cands
                    ]
                    escolha = st.session_state.get(f"prev_match_{ver}_{i}", IGNORAR)
                    idx = opcoes.index(escolha) - 1 if escolha in opcoes else -1
                    qtd_f = float(st.session_state.get(f"prev_qtd_{ver}_{i}", item["quantidade_gramas"]) or 0)
                    if idx < 0 or qtd_f <= 0:
                        ignorados.append(item["nome_original"])
                        continue
                    cand = cands[idx]
                    comp = db.get_composicao_por_100g(cand["fonte"], cand["id"])
                    if not comp:
                        ignorados.append(item["nome_original"])
                        continue
                    st.session_state.ingredientes.append(_novo_ingrediente(
                        nome=cand["nome"], fonte=cand["fonte"], fonte_id=cand["id"],
                        qtd_gramas=qtd_f, comp=comp,
                    ))
                    importados += 1

                st.session_state.upload_preview = None
                _flash(f"{importados} ingrediente(s) importado(s) para a receita.")
                if ignorados:
                    _flash(
                        "Itens do arquivo que NÃO entraram na receita (sem correspondência "
                        "selecionada ou quantidade zero): " + ", ".join(ignorados),
                        "warning",
                    )
                st.rerun()

    # ── Tabela de ingredientes adicionados ────────────────────────────────────
    sty.render_divider()
    sty.render_section_label(f"Ingredientes da Receita ({len(st.session_state.ingredientes)})")

    if not st.session_state.ingredientes:
        st.info("Nenhum ingrediente adicionado. Use a busca acima para começar.")
    else:
        ch1, ch2, ch3, ch4, ch5 = st.columns([4, 2, 2, 2, 1])
        ch1.markdown("**Ingrediente**")
        ch2.markdown("**Fonte**")
        ch3.markdown("**Qtd. original**")
        ch4.markdown("**Qtd. (g)**")
        ch5.markdown("**Rem.**")

        for ing in st.session_state.ingredientes:
            uid = ing["uid"]
            c1, c2, c3, c4, c5 = st.columns([4, 2, 2, 2, 1])
            c1.write(ing["nome"])
            c2.markdown(sty.render_fonte_badge(ing["fonte"]), unsafe_allow_html=True)

            unid_orig = ing.get("unidade_original") or "g"
            c3.caption(f"{calc.smart_round(ing.get('quantidade_original'))} {unid_orig}")

            nova_qtd = c4.number_input(
                f"Quantidade em gramas de {ing['nome']}",
                value=float(ing["quantidade_gramas"]),
                min_value=0.0,
                step=1.0,
                format="%.4f",
                key=f"qtd_ing_{uid}",
                label_visibility="collapsed",
            )
            if abs(nova_qtd - ing["quantidade_gramas"]) > 1e-9:
                ing["quantidade_gramas"] = nova_qtd
                # a quantidade passou a ser a digitada em gramas: o registro em
                # mL/% deixaria de corresponder ao que será calculado
                ing["unidade_original"]    = "g"
                ing["quantidade_original"] = nova_qtd
                ing["densidade_utilizada"] = 1.0
                st.rerun()

            c5.button("X", key=f"del_ing_{uid}", help="Remover ingrediente",
                      on_click=_cb_remover_ingrediente, args=(uid,))

        peso_total = sum(i["quantidade_gramas"] for i in st.session_state.ingredientes)
        pct_ings = [i for i in st.session_state.ingredientes if i.get("unidade_original") == "%"]
        if pct_ings:
            soma_pct = sum(i.get("quantidade_original") or 0 for i in pct_ings)
            pct_msg  = f"Soma %: {calc.smart_round(soma_pct)}% de {len(pct_ings)} ingrediente(s)"
            if abs(soma_pct - 100.0) < 0.01:
                st.success(pct_msg + " — 100% atingido")
            elif soma_pct > 100.0:
                st.error(pct_msg + f" — excede 100% em {soma_pct - 100:.4f}%")
            else:
                st.warning(pct_msg + f" — faltam {100 - soma_pct:.4f}%")

        st.caption(f"Peso total da receita: {calc.smart_round(peso_total)} g")

    # ── Botões de ação ────────────────────────────────────────────────────────
    sty.render_divider()
    col_calc, col_limpar = st.columns([3, 1])

    with col_calc:
        calcular_btn = st.button(
            "Calcular Tabela Nutricional",
            width="stretch",
            type="primary",
            key="btn_calcular",
        )

    with col_limpar:
        st.button("Limpar Ingredientes", width="stretch", key="btn_limpar",
                  on_click=_cb_limpar_ingredientes)

    if calcular_btn:
        erros = _validar_entrada(cabecalho)
        for msg in erros:
            st.error(msg)
        if not erros:
            with st.spinner("Calculando composição nutricional..."):
                try:
                    st.session_state.resultado_calculo = _calcular(cabecalho, registrar_auditoria=True)
                except Exception as e:
                    st.session_state.resultado_calculo = None
                    st.error(f"Erro no cálculo: {e}")
    elif st.session_state.resultado_calculo is not None:
        # Algum dado mudou depois do último cálculo → refaz, para a tabela em
        # tela, o Excel e a receita salva nunca ficarem com números antigos.
        if st.session_state.resultado_calculo.get("assinatura") != _assinatura_entrada(cabecalho):
            if _validar_entrada(cabecalho):
                st.session_state.resultado_calculo = None
                st.info("Os dados da receita mudaram. Complete-os e clique em Calcular novamente.")
            else:
                try:
                    st.session_state.resultado_calculo = _calcular(cabecalho, registrar_auditoria=False)
                except Exception as e:
                    st.session_state.resultado_calculo = None
                    st.error(f"Erro no cálculo: {e}")

    # ── Resultado do Cálculo ──────────────────────────────────────────────────
    if st.session_state.resultado_calculo:
        res    = st.session_state.resultado_calculo
        tabela = res["tabela"]
        por_100 = tabela["por_100g"]
        porcao  = tabela["por_porcao"]
        vd_     = tabela["vd"]
        porcao_txt = calc.formatar_rotulo(tabela["info"]["porcao_gramas"], "g")

        sty.render_divider()
        sty.render_section_label("Informação Nutricional")
        medida_txt = f" ({tabela['info']['medida_caseira']})" if tabela["info"]["medida_caseira"] else ""
        st.caption(
            f"Porção de {porcao_txt}{medida_txt}  |  "
            f"Porções por embalagem: {res['num_porcoes']}  |  "
            f"Peso total da receita: {calc.smart_round(res['peso_total'])} g"
        )

        for aviso in res.get("avisos", []):
            st.warning(aviso)

        LINHAS_ROTULO = [
            ("Valor Energético",       "energia_kcal",         "kcal"),
            ("",                       "energia_kj",           "kJ"),
            ("Carboidratos",           "carboidrato",          "g"),
            ("  Açúcares Totais",      "acucares_totais",      "g"),
            ("  Açúcares Adicionados", "acucares_adicionados", "g"),
            ("Proteínas",              "proteina",             "g"),
            ("Gorduras Totais",        "lipideos",             "g"),
            ("  Gorduras Saturadas",   "gordura_saturada",     "g"),
            ("  Gorduras Trans",       "gordura_trans",        "g"),
            ("Fibra Alimentar",        "fibra_alimentar",      "g"),
            ("Sódio",                  "sodio",                "mg"),
        ]

        dados_tabela = []
        for nome_nut, chave, unidade_nut in LINHAS_ROTULO:
            vd_val = vd_.get(chave)
            if chave == "energia_kj":
                vd_str = ""
            elif vd_val is None:
                vd_str = "**"
            else:
                vd_str = f"{vd_val}%"
            dados_tabela.append({
                "Nutriente":                 nome_nut,
                "Por 100 g":                 calc.formatar_rotulo(por_100.get(chave), unidade_nut),
                f"Por Porção ({porcao_txt})": calc.formatar_rotulo(porcao.get(chave), unidade_nut),
                "%VD*":                      vd_str,
            })

        # altura para todas as linhas: sem ela a última (Sódio) ficava escondida
        # atrás de uma barra de rolagem
        st.dataframe(pd.DataFrame(dados_tabela), hide_index=True, width="stretch",
                     height=35 * (len(dados_tabela) + 1) + 3)

        st.caption(
            "*Percentual de valores diários fornecidos pela porção. "
            "**Valor Diário não estabelecido."
        )
        if porcao.get("acucares_totais") is None:
            st.caption(
                "Açúcares totais: não informado — a TACO não traz este dado. "
                "Preencha em 'Campos Adicionais ANVISA' para que apareça no rótulo."
            )
        sty.render_legal_notice()

        # ── Rastreabilidade das Fontes ────────────────────────────────────────
        alertas_val = res.get("alertas_validacao", [])
        fontes_usadas = sorted({
            (d.get("composicao_100g") or {}).get("fonte_dados") or d.get("fonte", "?")
            for d in res["detalhes"]
        })

        with st.expander(
            "Rastreabilidade das Fontes"
            + (f"  ·  ⚠️ {len(alertas_val)} ingrediente(s) completado(s) com TBCA"
               if alertas_val else "  ·  " + ", ".join(fontes_usadas)),
            expanded=bool(alertas_val),
        ):
            for alerta in alertas_val:
                nivel = alerta.get("nivel", "")
                cor_nivel = (
                    "#3FB950" if "Alta" in nivel
                    else "#E3B341" if "moderada" in nivel
                    else "#F85149"
                )
                score_pct = int(alerta.get("score_match", 0) * 100)
                linhas_corr = "".join(
                    f"<br>• <strong>{escape(str(c['nutriente']))}</strong>: "
                    f"{c['val_taco']:.4g} (TACO) → {c['val_tbca']:.4g} (TBCA), por 100 g"
                    for c in alerta.get("corrigidos", [])
                )
                linhas_div = "".join(
                    f"<br>· {escape(str(d))}" for d in alerta.get("divergencias", [])
                )
                st.markdown(
                    f"""<div style="border:1px solid rgba(255,200,50,0.3);
                        border-left:3px solid #E3B341;border-radius:6px;
                        padding:0.65rem 1rem;margin-bottom:0.6rem;
                        background:rgba(255,200,50,0.06)">
                    <strong>⚠ {escape(str(alerta['ingrediente']))}</strong> — valores ausentes
                    na TACO preenchidos com a TBCA<br>
                    <span style="font-size:0.82rem;color:var(--text-muted)">
                      Motivo: {escape(str(alerta['motivo']))}<br>
                      Alimento na TBCA: <em>{escape(str(alerta['match_tbca'] or 'não encontrado'))}</em>
                      ({score_pct}% de similaridade no nome)<br>
                      Nível de confiança: <span style="color:{cor_nivel};font-weight:600">
                        {escape(nivel)}</span>
                      {linhas_corr}
                      {('<br>Diferenças entre as fontes (valor da TACO mantido):' + linhas_div) if linhas_div else ''}
                    </span></div>""",
                    unsafe_allow_html=True,
                )

            for det in res["detalhes"]:
                if not det.get("validacao", {}).get("triangulacao_aplicada"):
                    fonte_dados = (det.get("composicao_100g") or {}).get("fonte_dados") \
                        or det.get("fonte", "?")
                    st.markdown(
                        f'<span style="color:var(--text-muted);font-size:0.8rem">'
                        f'✔ <strong>{escape(str(det["nome"]))}</strong> — '
                        f'Fonte: {escape(str(fonte_dados))}</span>',
                        unsafe_allow_html=True,
                    )

        sty.render_divider()

        col_dl, col_salvar = st.columns(2)

        with col_dl:
            if "excel_bytes" not in res:
                try:
                    res["excel_bytes"] = xls.gerar_excel(
                        nome_produto=res["nome_produto"],
                        porcao_gramas=res["porcao_g"],
                        num_porcoes=res["num_porcoes"],
                        medida_caseira=res["medida_caseira"],
                        ingredientes_detalhes=res["detalhes"],
                        por_100g=tabela["por_100g"],
                        por_porcao=tabela["por_porcao"],
                        vd=vd_,
                        peso_total_gramas=res["peso_total"],
                        alertas_validacao=res.get("alertas_validacao", []),
                    )
                except Exception as e:
                    st.error(f"Erro ao gerar Excel: {e}")
                    res["excel_bytes"] = None

            if res.get("excel_bytes"):
                st.download_button(
                    label="Baixar Excel — Tabela Nutricional + P&D",
                    data=res["excel_bytes"],
                    file_name=f"NutriCalc_{_nome_arquivo_seguro(res['nome_produto'])}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch",
                    type="primary",
                    on_click="ignore",
                )

        with col_salvar:
            rotulo_salvar = (
                f"Atualizar Receita #{st.session_state.receita_editando_id}"
                if st.session_state.receita_editando_id else "Salvar Receita"
            )
            if st.button(rotulo_salvar, width="stretch", key="btn_salvar"):
                try:
                    rid = db.salvar_receita(
                        nome=res["nome_produto"],
                        porcao_gramas=res["porcao_g"],
                        ingredientes=[{
                            "nome":               i["nome"],
                            "fonte":              i["fonte"],
                            "fonte_id":           i.get("fonte_id"),
                            "quantidade_gramas":  i["quantidade_gramas"],
                            "unidade_original":   i.get("unidade_original", "g"),
                            "quantidade_original": i.get("quantidade_original"),
                            "densidade_utilizada": i.get("densidade_utilizada", 1.0),
                        } for i in st.session_state.ingredientes],
                        num_porcoes=res["num_porcoes"],
                        medida_caseira=res["medida_caseira"],
                        receita_id_existente=st.session_state.receita_editando_id,
                        extras=res["extras"],
                    )
                    st.session_state.receita_editando_id = rid
                    _flash(
                        f"Receita salva (ID #{rid}). Lembre-se: baixe o backup na aba "
                        "'Receitas Salvas' antes de encerrar — o servidor apaga os dados ao reiniciar."
                    )
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao salvar receita: {e}")


# ══════════════════════════════════════════════════════════════════════════════
#  ABA 2 — INGREDIENTES DE FORNECEDOR
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    sty.render_section_label("Cadastro de Ingredientes Industrializados")
    st.caption("Cadastre ingredientes não encontrados na TACO: aditivos, concentrados, extratos, etc.")

    NUTRIENTES_LABELS_FORM = {
        "umidade":        "Umidade (%)",
        "energia_kcal":   "Energia (kcal)",
        "energia_kj":     "Energia (kJ)",
        "proteina":       "Proteínas (g)",
        "lipideos":       "Gorduras Totais (g)",
        "colesterol":     "Colesterol (mg)",
        "carboidrato":    "Carboidratos (g)",
        "fibra_alimentar":"Fibra Alimentar (g)",
        "cinzas":         "Cinzas (g)",
        "calcio":         "Cálcio (mg)",
        "magnesio":       "Magnésio (mg)",
        "manganes":       "Manganês (mg)",
        "fosforo":        "Fósforo (mg)",
        "ferro":          "Ferro (mg)",
        "sodio":          "Sódio (mg)",
        "potassio":       "Potássio (mg)",
        "cobre":          "Cobre (mg)",
        "zinco":          "Zinco (mg)",
        "retinol":        "Retinol (mcg)",
        "re":             "RE (mcg)",
        "rae":            "RAE (mcg)",
        "tiamina":        "Tiamina (mg)",
        "riboflavina":    "Riboflavina (mg)",
        "piridoxina":     "Piridoxina (mg)",
        "niacina":        "Niacina (mg)",
        "vitamina_c":     "Vitamina C (mg)",
        "vitamina_d":     "Vitamina D (mcg)",
        "vitamina_e":     "Vitamina E (mg)",
        "vitamina_b12":   "Vitamina B12 (mcg)",
    }

    with st.expander("Pré-preencher formulário a partir de ficha técnica do fornecedor", expanded=False):
        st.caption(
            "Carregue um PDF, Excel ou CSV com a tabela nutricional por 100 g do ingrediente. "
            "Os campos serão preenchidos automaticamente."
        )
        arq_ficha_forn = st.file_uploader(
            "Selecionar ficha técnica",
            type=["pdf", "xlsx", "xls", "csv", "txt"],
            key="upload_ficha_forn",
            label_visibility="collapsed",
        )
        col_btn_ext, col_btn_clear = st.columns([2, 1])
        with col_btn_ext:
            if arq_ficha_forn and st.button("Extrair dados nutricionais", key="btn_extrair_nutri"):
                with st.spinner("Analisando ficha técnica..."):
                    try:
                        ext_arq   = arq_ficha_forn.name.rsplit(".", 1)[-1]
                        dados_ext = prs.parse_composicao_nutricional(arq_ficha_forn.getvalue(), ext_arq)
                        if "nome_comercial" not in dados_ext:
                            dados_ext["nome_comercial"] = (
                                arq_ficha_forn.name.rsplit(".", 1)[0]
                                .replace("_", " ").replace("-", " ").title()
                            )
                        st.session_state.forn_dados_extraidos = dados_ext
                        st.session_state.forn_editando_id     = None
                        st.session_state.forn_form_ver       += 1
                        n = len([
                            k for k in dados_ext
                            if k != "nome_comercial" and dados_ext[k] is not None
                        ])
                        _flash(f"{n} nutrientes extraídos. Confira o formulário pré-preenchido na aba "
                               "'Ingredientes de Fornecedor' antes de salvar.")
                        st.rerun()
                    except ValueError as e:
                        st.error(str(e))
                    except Exception as e:
                        st.error(f"Erro inesperado ao ler a ficha: {type(e).__name__}. "
                                 "Use o formulário manual.")
        with col_btn_clear:
            if st.session_state.forn_dados_extraidos:
                st.button("Limpar extração", key="btn_clear_ext",
                          on_click=_cb_cancelar_edicao_fornecedor)

        if st.session_state.forn_dados_extraidos:
            campos = [
                NUTRIENTES_LABELS_FORM.get(k, k)
                for k, v in st.session_state.forn_dados_extraidos.items()
                if k != "nome_comercial" and v is not None
            ]
            st.caption(f"Campos pré-preenchidos: {', '.join(campos)}")

    dados_edicao = {}
    if st.session_state.forn_editando_id:
        dados_edicao = db.get_ingrediente_fornecedor(st.session_state.forn_editando_id) or {}
        if not dados_edicao:                      # foi excluído enquanto era editado
            st.session_state.forn_editando_id = None
        else:
            ce1, ce2 = st.columns([4, 1])
            ce1.info(f"Editando: {dados_edicao.get('nome_comercial', '')}")
            ce2.button("Cancelar edição", key="btn_forn_cancelar", width="stretch",
                       on_click=_cb_cancelar_edicao_fornecedor)

    dados_form = dados_edicao if st.session_state.forn_editando_id else st.session_state.forn_dados_extraidos
    # A versão entra na chave de cada campo: quando os dados de origem mudam
    # (carregar para edição, extrair ficha, salvar), os campos são recriados com
    # os valores novos. Com chave fixa o Streamlit mantinha o que estava digitado.
    fv = st.session_state.forn_form_ver

    with st.form(f"form_fornecedor_{fv}", clear_on_submit=False):
        st.markdown("**Identificação**")
        cf1, cf2, cf3 = st.columns(3)
        f_nome_com   = cf1.text_input("Nome Comercial *",  value=dados_form.get("nome_comercial") or "")
        f_nome_gen   = cf2.text_input("Nome Genérico",     value=dados_form.get("nome_generico") or "")
        f_fabricante = cf3.text_input("Fabricante",        value=dados_form.get("fabricante") or "")
        f_obs        = st.text_area("Observações",         value=dados_form.get("observacoes") or "",
                                    height=68)

        st.markdown("**Valores Nutricionais por 100 g**")
        st.caption("Preencha com base na ficha técnica do fornecedor.")

        OBRIGATORIOS = ["energia_kcal", "energia_kj", "proteina", "lipideos",
                        "carboidrato",  "fibra_alimentar", "sodio"]

        st.markdown("*Obrigatórios ANVISA:*")
        cols_ob   = st.columns(len(OBRIGATORIOS))
        vals_forn = {}
        for i, nut in enumerate(OBRIGATORIOS):
            vals_forn[nut] = cols_ob[i].number_input(
                NUTRIENTES_LABELS_FORM.get(nut, nut),
                min_value=0.0,
                value=float(dados_form.get(nut) or 0.0),
                step=0.01,
                key=f"forn_{nut}_{fv}",
            )

        with st.expander("Micronutrientes adicionais"):
            extras  = [n for n in NUTRIENTES if n not in OBRIGATORIOS]
            cols_ex = st.columns(5)
            for i, nut in enumerate(extras):
                vals_forn[nut] = cols_ex[i % 5].number_input(
                    NUTRIENTES_LABELS_FORM.get(nut, nut),
                    min_value=0.0,
                    value=float(dados_form.get(nut) or 0.0),
                    step=0.001,
                    key=f"forn_{nut}_extra_{fv}",
                    format="%.4f",
                )

        label_btn = "Atualizar Ingrediente" if st.session_state.forn_editando_id else "Salvar Ingrediente"
        sub_forn  = st.form_submit_button(label_btn, width="stretch", type="primary")

    if sub_forn:
        if not f_nome_com.strip():
            st.error("O Nome Comercial é obrigatório.")
        else:
            try:
                kcal, kj = vals_forn.get("energia_kcal") or 0, vals_forn.get("energia_kj") or 0
                if kcal and not kj:
                    vals_forn["energia_kj"] = round(kcal * 4.184, 1)
                elif kj and not kcal:
                    vals_forn["energia_kcal"] = round(kj / 4.184, 1)
                dados = {
                    "nome_comercial": f_nome_com.strip(),
                    "nome_generico":  f_nome_gen.strip() or None,
                    "fabricante":     f_fabricante.strip() or None,
                    "observacoes":    f_obs.strip() or None,
                    **vals_forn,
                }
                if st.session_state.forn_editando_id:
                    fid = st.session_state.forn_editando_id
                    db.atualizar_ingrediente_fornecedor(fid, dados)
                    # receitas em edição passam a usar a composição atualizada
                    nova_comp = db.get_composicao_por_100g("FORNECEDOR", fid)
                    for ing in st.session_state.ingredientes:
                        if ing["fonte"] == "FORNECEDOR" and ing["fonte_id"] == fid and nova_comp:
                            ing["composicao_100g"] = nova_comp
                            ing["nome"]            = nova_comp["nome"]
                    st.session_state.resultado_calculo = None
                    _flash(f"Ingrediente '{f_nome_com.strip()}' atualizado.")
                    st.session_state.forn_editando_id = None
                else:
                    db.salvar_ingrediente_fornecedor(dados)
                    _flash(f"Ingrediente '{f_nome_com.strip()}' cadastrado com sucesso.")
                st.session_state.forn_dados_extraidos = {}
                st.session_state.forn_form_ver       += 1
                st.rerun()
            except Exception as e:
                st.error(f"Erro ao salvar ingrediente: {e}")

    sty.render_divider()
    sty.render_section_label("Ingredientes Cadastrados")

    try:
        forn_lista = db.listar_ingredientes_fornecedor()
        if not forn_lista:
            st.info("Nenhum ingrediente de fornecedor cadastrado ainda.")
        else:
            df_forn = pd.DataFrame(forn_lista)[[
                "id", "nome_comercial", "nome_generico", "fabricante",
                "energia_kcal", "proteina", "lipideos", "carboidrato", "sodio", "data_cadastro",
            ]].rename(columns={
                "id":             "ID",
                "nome_comercial": "Nome Comercial",
                "nome_generico":  "Nome Genérico",
                "fabricante":     "Fabricante",
                "energia_kcal":   "kcal",
                "proteina":       "Prot.(g)",
                "lipideos":       "Gord.(g)",
                "carboidrato":    "Carb.(g)",
                "sodio":          "Sódio(mg)",
                "data_cadastro":  "Cadastrado em",
            })
            st.dataframe(df_forn, hide_index=True, width="stretch")

            # rótulo com o ID: dois ingredientes com o mesmo nome continuam distinguíveis
            mapa_forn  = {f"#{f['id']} — {f['nome_comercial']}": f["id"] for f in forn_lista}
            nomes_forn = list(mapa_forn)

            c_edit, c_del = st.columns(2)
            with c_edit:
                sel_editar = st.selectbox(
                    "Editar ingrediente:", [SEM_SELECAO] + nomes_forn, key="sel_forn_editar"
                )
                if sel_editar != SEM_SELECAO:
                    st.button("Carregar para Edição", key="btn_forn_edit",
                              on_click=_cb_editar_fornecedor, args=(mapa_forn,))

            with c_del:
                sel_deletar = st.selectbox(
                    "Excluir ingrediente:", [SEM_SELECAO] + nomes_forn, key="sel_forn_del"
                )
                if sel_deletar != SEM_SELECAO:
                    st.caption("Receitas salvas que usam este ingrediente ficarão sem ele.")
                    if st.button("Excluir", key="btn_forn_del", type="secondary"):
                        fid_del = mapa_forn[sel_deletar]
                        db.deletar_ingrediente_fornecedor(fid_del)
                        if st.session_state.forn_editando_id == fid_del:
                            st.session_state.forn_editando_id = None
                            st.session_state.forn_form_ver   += 1
                        _flash(f"'{sel_deletar}' excluído.")
                        st.rerun()
    except Exception as e:
        st.error(f"Erro ao carregar ingredientes: {e}")


# ══════════════════════════════════════════════════════════════════════════════
#  ABA 3 — RECEITAS SALVAS
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    sty.render_section_label("Receitas Salvas")

    try:
        receitas = db.listar_receitas()
        if not receitas:
            st.info(
                "Nenhuma receita salva ainda. "
                "Calcule uma receita na aba 'Nova Receita' e clique em 'Salvar Receita'."
            )
        else:
            df_rec = pd.DataFrame(receitas)[[
                "id", "nome_produto", "porcao_gramas", "num_porcoes",
                "num_ingredientes", "data_criacao",
            ]].rename(columns={
                "id":               "ID",
                "nome_produto":     "Produto",
                "porcao_gramas":    "Porção (g)",
                "num_porcoes":      "Nº Porções",
                "num_ingredientes": "Ingredientes",
                "data_criacao":     "Criado em",
            })
            st.dataframe(df_rec, hide_index=True, width="stretch")

            sty.render_divider()
            mapa_rec  = {f"#{r['id']} — {r['nome_produto']}": r["id"] for r in receitas}
            nomes_rec = list(mapa_rec)
            c1, c2, c3 = st.columns(3)

            with c1:
                st.markdown("**Abrir / Reabrir Receita**")
                sel_abrir = st.selectbox("Selecionar:", [SEM_SELECAO] + nomes_rec, key="sel_abrir")
                if sel_abrir != SEM_SELECAO:
                    if st.session_state.ingredientes and not st.session_state.receita_editando_id:
                        st.caption("Atenção: a receita em edição na aba 'Nova Receita' "
                                   "ainda não foi salva e será substituída.")
                    st.button("Abrir para Edição", width="stretch", key="btn_abrir",
                              on_click=_cb_abrir_receita, args=(mapa_rec,))

            with c2:
                st.markdown("**Duplicar Receita**")
                sel_dup = st.selectbox("Selecionar:", [SEM_SELECAO] + nomes_rec, key="sel_dup")
                novo_nome_dup = st.text_input(
                    "Nome da cópia:", key="nome_dup", placeholder="Ex.: Bolo Chocolate Diet v2"
                )
                if sel_dup != SEM_SELECAO and novo_nome_dup.strip():
                    if st.button("Duplicar", width="stretch", key="btn_dup"):
                        try:
                            novo_id = db.duplicar_receita(mapa_rec[sel_dup], novo_nome_dup.strip())
                            _flash(f"Receita duplicada como '{novo_nome_dup.strip()}' (ID #{novo_id}).")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Erro ao duplicar: {e}")

            with c3:
                st.markdown("**Excluir Receita**")
                sel_del = st.selectbox("Selecionar:", [SEM_SELECAO] + nomes_rec, key="sel_del")
                if sel_del != SEM_SELECAO:
                    if st.button(
                        "Excluir Receita", width="stretch",
                        key="btn_del_rec", type="secondary",
                    ):
                        rid_del = mapa_rec[sel_del]
                        db.deletar_receita(rid_del)
                        if st.session_state.receita_editando_id == rid_del:
                            st.session_state.receita_editando_id = None
                        _flash(f"Receita '{sel_del}' excluída.")
                        st.rerun()

        sty.render_divider()
        st.markdown("#### Backup de Dados")
        st.caption(
            "O servidor reinicia a cada ~1 hora sem uso e os dados são apagados. "
            "Baixe um backup antes de encerrar a sessão e importe-o ao voltar."
        )
        col_exp, col_imp = st.columns(2)

        with col_exp:
            st.markdown("**Exportar Backup**")
            n_forn_bkp = len(db.listar_ingredientes_fornecedor())
            if not receitas and not n_forn_bkp:
                st.caption("Ainda não há receitas nem ingredientes de fornecedor para exportar.")
            else:
                # Gerado na hora do clique: o arquivo sempre reflete o estado atual
                # (antes o backup ficava guardado na sessão e saía desatualizado).
                st.download_button(
                    label=f"⬇️ Baixar backup — {len(receitas)} receita(s), {n_forn_bkp} ingrediente(s)",
                    data=db.export_backup,
                    file_name=f"nutricalc_backup_{datetime.now():%Y-%m-%d_%Hh%M}.json",
                    mime="application/json",
                    width="stretch",
                    key="btn_dl_backup",
                    on_click="ignore",
                )

        with col_imp:
            st.markdown("**Importar Backup**")
            arq_backup = st.file_uploader(
                "Selecionar arquivo de backup",
                type=["json"],
                key="upload_backup",
                label_visibility="collapsed",
            )
            if arq_backup is not None:
                if st.button("Importar", width="stretch", key="btn_importar_backup"):
                    try:
                        r = db.import_backup(arq_backup.getvalue())
                        msg = (f"Backup importado: {r['receitas']} receita(s) e "
                               f"{r['ingredientes']} ingrediente(s) de fornecedor restaurados.")
                        if r["receitas_ignoradas"] or r["ingredientes_ignorados"]:
                            msg += (f" Já existiam e foram mantidos: {r['receitas_ignoradas']} "
                                    f"receita(s), {r['ingredientes_ignorados']} ingrediente(s).")
                        _flash(msg)
                        st.rerun()
                    except Exception as ex:
                        st.error(f"Erro ao importar backup: {ex}")

    except Exception as e:
        st.error(f"Erro ao carregar receitas: {e}")
