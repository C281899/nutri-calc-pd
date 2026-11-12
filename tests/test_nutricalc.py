"""
Testes de regressão do NutriCalc P&D.

    python -m pytest tests -q

Cobrem as regras da IN 75/2020 (arredondamento, %VD, quantidades não
significativas), a leitura de números em formato brasileiro, o ciclo
salvar → backup → restaurar e o fluxo principal da interface.
Usam bancos temporários: não tocam nos dados reais.
"""
import io
import json
import os
import shutil
import sys

import pytest

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if RAIZ not in sys.path:
    sys.path.insert(0, RAIZ)

from modules import calculator as calc          # noqa: E402
from modules import parser as prs               # noqa: E402


# ─── Arredondamento (IN 75/2020, Anexo III) ───────────────────────────────────

@pytest.mark.parametrize("valor, unidade, esperado", [
    (123.5, "kcal", 124), (123.4, "kcal", 123), (10.5, "g", 11),
    (9.95, "g", 10), (9.94, "g", 9.9), (2.25, "g", 2.3), (2.24, "g", 2.2),
    (0.25, "g", 0.3), (0.04, "g", 0.0), (0.05, "g", 0.1),
    (0.125, "mg", 0.13), (0.124, "mg", 0.12), (4.6, "mg", 4.6), (148.5, "mg", 149),
    (0, "g", 0), (-1, "g", 0), (None, "g", 0),
])
def test_arredondamento_anvisa(valor, unidade, esperado):
    assert calc.arredondar_anvisa(valor, unidade) == pytest.approx(esperado)


@pytest.mark.parametrize("valor, unidade, texto", [
    (124.0, "kcal", "124 kcal"), (2.3, "g", "2,3 g"), (2.0, "g", "2 g"),
    (0.13, "mg", "0,13 mg"), (0.1, "mg", "0,1 mg"), (0.0, "g", "0 g"),
    (None, "g", "Não informado"),
])
def test_formatacao_rotulo(valor, unidade, texto):
    assert calc.formatar_rotulo(valor, unidade) == texto


# ─── Valores Diários (IN 75/2020, Anexo II) ───────────────────────────────────

def test_valores_diarios_de_referencia():
    assert calc.VALORES_DIARIOS == {
        "energia_kcal": 2000, "energia_kj": 8400, "carboidrato": 300,
        "acucares_adicionados": 50, "proteina": 50, "lipideos": 65,
        "gordura_saturada": 20, "gordura_trans": 2, "fibra_alimentar": 25,
        "sodio": 2000,
    }


def test_percentual_vd():
    vd = calc.calcular_vd({
        "energia_kcal": 150, "carboidrato": 30, "acucares_adicionados": 5,
        "proteina": 5, "lipideos": 6.5, "gordura_saturada": 2,
        "gordura_trans": 0.2, "fibra_alimentar": 2.5, "sodio": 100,
    })
    assert vd["energia_kcal"] == 8          # 7,5 → 8 (meio para cima)
    assert vd["carboidrato"] == 10
    assert vd["proteina"] == 10
    assert vd["gordura_saturada"] == 10
    assert vd["gordura_trans"] == 10
    assert vd["sodio"] == 5
    assert vd["acucares_totais"] is None


# ─── Tabela do rótulo ─────────────────────────────────────────────────────────

def _tabela(totais, peso=1000.0, porcao=30.0, **extras):
    return calc.montar_tabela_rotulo("Teste", porcao, peso, totais, **extras)


def test_tabela_por_100g_e_por_porcao():
    t = _tabela({"energia_kcal": 4000, "energia_kj": 16736, "carboidrato": 600,
                 "proteina": 100, "lipideos": 150, "fibra_alimentar": 30, "sodio": 2500},
                gordura_saturada_total=50, acucares_adicionados_total=200,
                acucares_totais_total=250)
    assert t["por_100g"]["energia_kcal"] == 400
    assert t["por_porcao"]["energia_kcal"] == 120
    assert t["por_porcao"]["carboidrato"] == 18
    assert t["por_porcao"]["proteina"] == 3
    assert t["por_porcao"]["gordura_saturada"] == 1.5
    assert t["por_porcao"]["acucares_totais"] == 7.5
    assert t["por_porcao"]["sodio"] == 75
    assert t["vd"]["sodio"] == 4            # 75 / 2000
    assert t["vd"]["proteina"] == 6         # 3 / 50


def test_acucares_totais_nao_informados_nao_sao_estimados():
    t = _tabela({"energia_kcal": 4000, "carboidrato": 600})
    assert t["por_100g"]["acucares_totais"] is None
    assert t["por_porcao"]["acucares_totais"] is None


def test_quantidade_nao_significativa_vira_zero_so_se_valer_para_100g_e_porcao():
    # sódio: 4 mg/100 g e 1,2 mg/porção → ambos ≤ 5 mg → zero
    t = _tabela({"sodio": 40})
    assert t["por_100g"]["sodio"] == 0 and t["por_porcao"]["sodio"] == 0
    # sódio: 12 mg/100 g e 3,6 mg/porção → 100 g é significativo → mantém os dois
    t = _tabela({"sodio": 120})
    assert t["por_100g"]["sodio"] == 12 and t["por_porcao"]["sodio"] == 3.6


def test_gordura_total_nao_zera_se_saturada_for_significativa():
    # 0,4 g de gordura/100 g (≤ 0,5), mas 0,3 g de saturada (> 0,1): não pode zerar
    t = _tabela({"lipideos": 4.0}, porcao=100.0, gordura_saturada_total=3.0)
    assert t["por_100g"]["lipideos"] == 0.4
    assert t["por_100g"]["gordura_saturada"] == 0.3


def test_carboidrato_nao_zera_se_acucar_for_significativo():
    t = _tabela({"carboidrato": 5.0}, porcao=100.0, acucares_totais_total=5.0 + 1.0)
    assert t["por_100g"]["carboidrato"] == 0.5


def test_peso_total_zero_nao_quebra():
    t = _tabela({"energia_kcal": 100}, peso=0.0)
    assert t["por_100g"]["energia_kcal"] == 0


# ─── Números em formato brasileiro ────────────────────────────────────────────

@pytest.mark.parametrize("texto, valor", [
    ("200,5", 200.5), ("1.250,00", 1250.0), ("2.5", 2.5), ("1.250", 1250.0),
    ("1,250.50", 1250.5), (" 30 g", 30.0), ("0,75", 0.75),
])
def test_numero_formato_brasileiro(texto, valor):
    assert prs._num_br(texto) == pytest.approx(valor)


def test_csv_de_receita_com_virgula_decimal():
    csv = "Ingrediente;Quantidade (g)\nFarinha de trigo;1.250,5\nAçúcar;300\nSal;2,5\n"
    itens = prs.parse_ficha_tecnica(csv.encode("utf-8"), "csv")
    por_nome = {i["nome"].lower(): i["quantidade_gramas"] for i in itens}
    assert por_nome["farinha de trigo"] == pytest.approx(1250.5)
    assert por_nome["sal"] == pytest.approx(2.5)


def test_texto_com_milhar_e_decimal_nao_perde_ingrediente():
    txt = "Ficha de produção\nFarinha de trigo: 1.250,5 g\nAçúcar refinado: 300 g\nSal: 2,5 g\nÁgua: 0,7 L\n"
    itens = prs.parse_ficha_tecnica(txt.encode("utf-8"), "txt")
    por_nome = {i["nome"].lower(): i["quantidade_gramas"] for i in itens}
    assert por_nome == pytest.approx(
        {"farinha de trigo": 1250.5, "açúcar refinado": 300.0, "sal": 2.5, "água": 700.0})


def test_energia_em_kj_e_completada_quando_a_ficha_so_traz_kcal():
    r = calc.calcular_composicao_receita(
        [{"nome": "X", "fonte": "FORNECEDOR", "quantidade_gramas": 100.0,
          "composicao_100g": {"energia_kcal": 100.0, "energia_kj": None}}], validar_taco=False)
    assert r["nutrientes"]["energia_kj"] == pytest.approx(418.4)


def test_ficha_de_um_ingrediente_nao_vira_receita():
    txt = ("INFORMAÇÃO NUTRICIONAL por 100 g\n"
           "Valor energético 350 kcal = 1470 kJ\nCarboidratos 80 g\nProteínas 6,5 g\n"
           "Gorduras totais 1,2 g\nGorduras saturadas 0,4 g\nFibra alimentar 3 g\nSódio 0,5 g\n")
    itens = prs.parse_ficha_tecnica(txt.encode("utf-8"), "txt")
    assert len(itens) == 1 and "composicao_direta" in itens[0]
    comp = itens[0]["composicao_direta"]
    assert comp["energia_kcal"] == 350
    assert comp["proteina"] == pytest.approx(6.5)
    assert comp["lipideos"] == pytest.approx(1.2)     # não confunde com saturadas
    assert comp["sodio"] == pytest.approx(500)        # 0,5 g → 500 mg


# ─── Banco: salvar, backup e restauração ──────────────────────────────────────

@pytest.fixture()
def bancos(tmp_path, monkeypatch):
    """Bancos novos em pasta temporária, com os módulos apontando para eles."""
    import importlib.util
    spec = importlib.util.spec_from_file_location(
        "taco_setup_teste", os.path.join(RAIZ, "data", "taco_setup.py"))
    setup = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(setup)
    monkeypatch.setattr(setup, "TACO_DB", str(tmp_path / "taco.db"))
    monkeypatch.setattr(setup, "APP_DB", str(tmp_path / "nutri_calc.db"))
    monkeypatch.setattr(setup, "TACO_XLS", str(tmp_path / "nao_existe.xlsx"))
    monkeypatch.setattr(setup, "baixar_taco_excel", lambda *a, **k: False)
    setup.criar_taco_db()
    setup.criar_app_db()

    from modules import database as db
    monkeypatch.setattr(db, "TACO_DB", str(tmp_path / "taco.db"))
    monkeypatch.setattr(db, "APP_DB", str(tmp_path / "nutri_calc.db"))
    monkeypatch.setattr(db, "_migrado_para", None)
    db._indice_taco.update(mtime=None, itens=[])
    db.migrar_banco()
    yield db
    db._indice_taco.update(mtime=None, itens=[])


def test_busca_prioriza_nome_mais_proximo(bancos):
    db = bancos
    nomes = [c["nome"] for c in db.buscar_ingrediente("farinha de trigo", 5)]
    assert nomes and "farinha" in nomes[0].lower() and "trigo" in nomes[0].lower()
    assert db.buscar_ingrediente("", 5) == []


def test_backup_restaura_e_nao_duplica(bancos):
    db = bancos
    fid = db.salvar_ingrediente_fornecedor(
        {"nome_comercial": "Aroma X", "fabricante": "ACME", "energia_kcal": 10.0})
    farinha = db.buscar_ingrediente("farinha de trigo", 1)[0]
    rid = db.salvar_receita(
        "Bolo", 30.0,
        [{"nome": farinha["nome"], "fonte": "TACO", "fonte_id": farinha["id"],
          "quantidade_gramas": 500.0},
         {"nome": "Aroma X", "fonte": "FORNECEDOR", "fonte_id": fid,
          "quantidade_gramas": 5.0}],
        num_porcoes=10, medida_caseira="1 fatia",
        extras={"gordura_saturada_total": 12.5, "acucares_totais_total": None},
    )
    backup = db.export_backup()

    db.deletar_receita(rid)
    db.deletar_ingrediente_fornecedor(fid)
    assert db.listar_receitas() == []

    r1 = db.import_backup(backup)
    assert (r1["receitas"], r1["ingredientes"]) == (1, 1)
    r2 = db.import_backup(backup)                       # importar de novo não duplica
    assert (r2["receitas"], r2["ingredientes"]) == (0, 0)
    assert len(db.listar_receitas()) == 1

    receita = db.get_receita_completa(db.listar_receitas()[0]["id"])
    assert receita["gordura_saturada_total"] == 12.5
    assert receita["acucares_totais_total"] is None
    aroma = next(i for i in receita["ingredientes"] if i["nome_ingrediente"] == "Aroma X")
    achado = db.resolver_ingrediente(aroma["fonte"], aroma["fonte_id"], aroma["nome_ingrediente"])
    assert achado and achado["nome"] == "Aroma X"       # religado ao id novo


def test_id_trocado_nao_abre_outro_alimento(bancos):
    db = bancos
    farinha = db.buscar_ingrediente("farinha de trigo", 1)[0]
    # id errado de propósito: tem de achar pelo nome, nunca devolver o alimento do id
    achado = db.resolver_ingrediente("TACO", farinha["id"] + 1, farinha["nome"])
    assert achado and achado["nome"] == farinha["nome"]
    assert db.resolver_ingrediente("TACO", 1, "Ingrediente que não existe") is None


def test_backup_malformado_da_erro_claro(bancos):
    with pytest.raises(ValueError):
        bancos.import_backup(b"isto nao e json")
    with pytest.raises(ValueError):
        bancos.import_backup(json.dumps({"qualquer": 1}).encode())


# ─── Excel ────────────────────────────────────────────────────────────────────

def test_excel_e_gerado_e_abre(bancos):
    db = bancos
    from modules import excel_generator as xls
    from openpyxl import load_workbook
    farinha = db.buscar_ingrediente("farinha de trigo", 1)[0]
    comp = db.get_composicao_por_100g("TACO", farinha["id"])
    bruto = calc.calcular_composicao_receita(
        [{"nome": farinha["nome"], "fonte": "TACO", "fonte_id": farinha["id"],
          "quantidade_gramas": 500.0, "composicao_100g": comp}], validar_taco=False)
    t = calc.montar_tabela_rotulo("Pão", 50.0, bruto["peso_total_gramas"], bruto["nutrientes"])
    dados = xls.gerar_excel(
        nome_produto="Pão", porcao_gramas=50.0, num_porcoes=10, medida_caseira="1 unidade",
        ingredientes_detalhes=calc.calcular_contribuicoes_percentuais(bruto["ingredientes_detalhes"]),
        por_100g=t["por_100g"], por_porcao=t["por_porcao"], vd=t["vd"],
        peso_total_gramas=bruto["peso_total_gramas"], alertas_validacao=[])
    wb = load_workbook(io.BytesIO(dados))
    assert len(wb.sheetnames) == 2
    textos = [str(c.value) for row in wb.worksheets[0].iter_rows() for c in row if c.value]
    assert any("Porção de 50 g (1 unidade)" in t_ for t_ in textos)
    assert any("Não informado" in t_ for t_ in textos)      # açúcares totais


# ─── Interface (fluxo principal) ──────────────────────────────────────────────

@pytest.fixture()
def app(tmp_path):
    """Roda o app numa cópia do projeto, para não mexer nos bancos reais."""
    from streamlit.testing.v1 import AppTest
    import streamlit as st
    copia = tmp_path / "app"
    shutil.copytree(RAIZ, copia, ignore=shutil.ignore_patterns(
        "*.db", "*.db-*", "__pycache__", ".git", "tests", "*.xlsx"))
    # sem rede nos testes: usa a base embutida
    setup = copia / "data" / "taco_setup.py"
    setup.write_text(setup.read_text(encoding="utf-8").replace(
        'URL_TACO = "', 'URL_TACO = "http://127.0.0.1:9/'), encoding="utf-8")
    antigos = {k: v for k, v in sys.modules.items() if k == "modules" or k.startswith("modules.")}
    for k in antigos:
        del sys.modules[k]
    sys.path.insert(0, str(copia))
    st.cache_resource.clear()
    at = AppTest.from_file(str(copia / "app.py"), default_timeout=120).run()
    assert not at.exception, at.exception
    yield at
    sys.path.remove(str(copia))
    for k in [k for k in sys.modules if k == "modules" or k.startswith("modules.")]:
        del sys.modules[k]
    sys.modules.update(antigos)
    st.cache_resource.clear()


def _botao(at, chave):
    return next(b for b in at.button if b.key == chave)


def _adicionar(at, termo, qtd):
    at.text_input(key="busca_ingrediente").set_value(termo).run()
    at.number_input(key="qtd_manual").set_value(qtd).run()
    _botao(at, "btn_adicionar").click().run()


def test_fluxo_completo_da_interface(app):
    at = app
    _adicionar(at, "farinha de trigo", 100.0)
    _adicionar(at, "açúcar refinado", 200.0)
    _adicionar(at, "leite em pó integral", 300.0)
    ings = at.session_state.ingredientes
    assert [i["quantidade_gramas"] for i in ings] == [100.0, 200.0, 300.0]

    # remover o primeiro não pode embaralhar as quantidades dos outros
    _botao(at, f"del_ing_{ings[0]['uid']}").click().run()
    assert [i["quantidade_gramas"] for i in at.session_state.ingredientes] == [200.0, 300.0]

    # calcular sem nome → erro, sem resultado
    _botao(at, "btn_calcular").click().run()
    assert at.session_state.resultado_calculo is None and at.error

    at.text_input(key="form_nome").set_value("Bolo Teste").run()
    _botao(at, "btn_calcular").click().run()
    assert not at.exception
    res = at.session_state.resultado_calculo
    assert res and res["porcao_g"] == 30.0 and res["peso_total"] == 500.0

    # mudar a porção depois de calcular refaz o resultado (nada de número antigo)
    at.number_input(key="form_porcao").set_value(50.0).run()
    assert at.session_state.resultado_calculo["porcao_g"] == 50.0

    _botao(at, "btn_salvar").click().run()
    assert not at.exception
    assert at.session_state.receita_editando_id is not None
    assert any("Receita salva" in s.value for s in at.success)

    # nova receita limpa tudo
    _botao(at, "btn_nova_receita").click().run()
    assert at.session_state.ingredientes == [] and at.session_state.form_nome == ""

    # reabrir a receita salva traz ingredientes e cabeçalho de volta
    opcao = next(o for o in at.selectbox(key="sel_abrir").options if "Bolo Teste" in o)
    at.selectbox(key="sel_abrir").set_value(opcao).run()
    _botao(at, "btn_abrir").click().run()
    assert not at.exception
    assert at.session_state.form_nome == "Bolo Teste"
    assert at.session_state.form_porcao == 50.0
    assert [i["quantidade_gramas"] for i in at.session_state.ingredientes] == [200.0, 300.0]
