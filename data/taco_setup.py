"""
taco_setup.py — Inicialização dos bancos de dados NutriCalc P&D
Executa UMA VEZ na primeira inicialização do sistema.

Uso:
    python data/taco_setup.py

O que faz:
    1. Cria taco.db com dados da TACO 4ª edição (UNICAMP)
    2. Cria nutri_calc.db com tabelas de receitas e ingredientes de fornecedor
"""

import re
import sqlite3
import os
import sys
import pandas as pd

# ─── Caminhos ──────────────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(BASE_DIR)
TACO_DB  = os.path.join(ROOT_DIR, "taco.db")
APP_DB   = os.path.join(ROOT_DIR, "nutri_calc.db")
TACO_XLS = os.path.join(BASE_DIR, "TACO_4ed.xlsx")

# URL oficial NEPA/UNICAMP para download automático
URL_TACO = "https://nepa.unicamp.br/wp-content/uploads/sites/27/2023/10/Taco-4a-Edicao.xlsx"

# ─── Mapeamento por posição de coluna (Excel oficial UNICAMP) ──────────────────
# O arquivo TACO tem cabeçalho mesclado em 3 linhas; usamos índice de coluna.
# Col 0 = ID, Col 1 = nome, Col 13 = ID repetido (ignorar).
# Vitamina D, E e B12 não constam no arquivo oficial — ficarão como NULL.
COL_NOME = 1
COL_NUTRIENTES = {
    'umidade': 2, 'energia_kcal': 3, 'energia_kj': 4,
    'proteina': 5, 'lipideos': 6, 'colesterol': 7,
    'carboidrato': 8, 'fibra_alimentar': 9, 'cinzas': 10,
    'calcio': 11, 'magnesio': 12,
    # col 13 = ID duplicado → ignorar
    'manganes': 14, 'fosforo': 15, 'ferro': 16, 'sodio': 17,
    'potassio': 18, 'cobre': 19, 'zinco': 20,
    'retinol': 21, 're': 22, 'rae': 23,
    'tiamina': 24, 'riboflavina': 25, 'piridoxina': 26,
    'niacina': 27, 'vitamina_c': 28,
}

# Mantido para compatibilidade com outros usos no código
COLUNAS_TACO = [
    ("nome_alimento",  []),
    ("categoria",      []),
    ("umidade",        []),
    ("energia_kcal",   []),
    ("energia_kj",     []),
    ("proteina",       []),
    ("lipideos",       []),
    ("colesterol",     []),
    ("carboidrato",    []),
    ("fibra_alimentar",[]),
    ("cinzas",         []),
    ("calcio",         []),
    ("magnesio",       []),
    ("manganes",       []),
    ("fosforo",        []),
    ("ferro",          []),
    ("sodio",          []),
    ("potassio",       []),
    ("cobre",          []),
    ("zinco",          []),
    ("retinol",        []),
    ("re",             []),
    ("rae",            []),
    ("tiamina",        []),
    ("riboflavina",    []),
    ("piridoxina",     []),
    ("niacina",        []),
    ("vitamina_c",     []),
    ("vitamina_d",     []),
    ("vitamina_e",     []),
    ("vitamina_b12",   []),
]

NUTRIENTES_COLUNAS = [col for col, _ in COLUNAS_TACO if col not in ("nome_alimento", "categoria")]

# ─── DDL dos bancos ─────────────────────────────────────────────────────────────
NUTRIENTES_DDL = "\n    ".join([f"{col} REAL," for col in NUTRIENTES_COLUNAS])

SQL_TACO_DB = f"""
CREATE TABLE IF NOT EXISTS alimentos_taco (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nome_alimento TEXT NOT NULL,
    categoria TEXT,
    {NUTRIENTES_DDL}
    fonte TEXT DEFAULT 'TACO 4ª edição'
);

CREATE INDEX IF NOT EXISTS idx_nome_taco ON alimentos_taco(nome_alimento);
"""

SQL_APP_DB = f"""
CREATE TABLE IF NOT EXISTS ingredientes_fornecedor (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nome_comercial TEXT NOT NULL,
    nome_generico TEXT,
    fabricante TEXT,
    {NUTRIENTES_DDL}
    data_cadastro TEXT DEFAULT (datetime('now','localtime')),
    observacoes TEXT
);

CREATE TABLE IF NOT EXISTS receitas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nome_produto TEXT NOT NULL,
    porcao_gramas REAL NOT NULL,
    num_porcoes INTEGER,
    medida_caseira TEXT,
    observacoes TEXT,
    data_criacao TEXT DEFAULT (datetime('now','localtime')),
    data_atualizacao TEXT
);

CREATE TABLE IF NOT EXISTS receita_ingredientes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    receita_id INTEGER NOT NULL REFERENCES receitas(id) ON DELETE CASCADE,
    nome_ingrediente TEXT NOT NULL,
    fonte TEXT NOT NULL,
    fonte_id INTEGER,
    quantidade_gramas REAL NOT NULL,
    eh_subrecita INTEGER DEFAULT 0,
    subrecita_id INTEGER REFERENCES receitas(id),
    ordem INTEGER DEFAULT 0,
    unidade_original TEXT DEFAULT 'g',
    quantidade_original REAL DEFAULT NULL,
    densidade_utilizada REAL DEFAULT 1.0
);

CREATE INDEX IF NOT EXISTS idx_receita_id ON receita_ingredientes(receita_id);

CREATE TABLE IF NOT EXISTS densidades_ingredientes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nome_ingrediente TEXT NOT NULL UNIQUE,
    densidade_g_ml REAL NOT NULL,
    fonte TEXT DEFAULT 'manual',
    data_cadastro TEXT DEFAULT (datetime('now','localtime'))
);
"""

# ─── Dados TACO embutidos (subconjunto mínimo para funcionar sem arquivo) ─────
# 50 alimentos mais comuns para permitir uso imediato, mesmo sem TACO_4ed.xlsx
TACO_EMBUTIDA = [
    # (nome, categoria, umidade, kcal, kj, prot, lip, col, carb, fibra, cinzas,
    #  ca, mg, mn, p, fe, na, k, cu, zn, retinol, re, rae, tia, ribo, pirid, niac, vitc, vitd, vite, vitb12)
    ("Açúcar refinado","Açúcares e produtos",0.1,387,1619,0,0,0,99.9,0,0.01,1,0,None,0,0.1,1,0,None,0,0,0,0,0,0,0,0,0,0,0,0),
    ("Sal refinado","Condimentos",0.1,0,0,0,0,0,0,0,99.9,None,None,None,None,None,38758,None,None,None,0,0,0,0,0,0,0,0,None,None,0),
    ("Farinha de trigo","Cereais e derivados",13.4,360,1506,9.8,1.4,0,75.1,2.3,0.3,18,25,0.7,126,1,2,107,0.12,0.81,0,0,0,0.17,0.04,0.05,1.5,0,None,0.08,0),
    ("Leite integral","Leite e derivados",87.9,61,255,3.2,3.2,10,4.8,0,0.7,123,12,0.01,90,0.04,38,163,0.01,0.39,46,46,14,0.04,0.18,0.05,0.09,0.9,None,0.06,0.45),
    ("Ovo de galinha inteiro","Ovos",75.9,143,598,13.3,9.5,396,0.6,0,1,59,13,0.04,222,1.88,146,147,0.07,1.33,174,174,174,0.06,0.46,0.05,0.08,0,None,1.01,1.4),
    ("Manteiga com sal","Óleos e gorduras",15.9,726,3039,0.4,80.5,219,0.1,0,2.1,11,3,0,15,0.31,714,26,0,0.12,672,672,174,0,0.01,0,0.01,0,None,2.35,0.17),
    ("Óleo de soja","Óleos e gorduras",0,884,3699,0,100,0,0,0,0,0,0,None,0,0,0,0,None,None,0,0,0,0,0,0,0,0,None,13.56,0),
    ("Banana prata","Frutas",74.5,98,410,1.3,0.1,0,23.8,2,0.3,3,27,0.35,22,0.29,1,366,0.11,0.16,7,7,4,0.03,0.05,0.37,0.71,9.1,None,0.13,0),
    ("Maçã fuji","Frutas",84.8,56,234,0.3,0.4,0,14.3,1.3,0.2,4,4,0.03,7,0.12,0,95,0.02,0.04,3,3,3,0.02,0.01,0.03,0.08,4.7,None,0.18,0),
    ("Tomate","Hortaliças",94.1,15,63,1.1,0.2,0,3.1,1.2,0.5,9,13,0.14,29,0.4,4,287,0.06,0.12,97,97,58,0.06,0.04,0.08,0.73,21.3,None,0.54,0),
    ("Cenoura","Hortaliças",88.6,40,167,0.9,0.2,0,9.3,3.2,0.8,28,13,0.2,29,0.4,44,305,0.08,0.18,2813,2813,1358,0.07,0.06,0.15,1.03,4.6,None,0.45,0),
    ("Batata inglesa","Hortaliças",80.2,66,276,1.7,0.1,0,15.5,1.8,0.8,5,22,0.17,44,0.33,6,401,0.14,0.3,0,0,0,0.09,0.03,0.3,1.44,17.1,None,0.01,0),
    ("Arroz polido","Cereais e derivados",12.4,360,1507,7.4,0.9,0,79,1.6,0.3,4,27,1.26,115,0.2,4,75,0.19,0.55,0,0,0,0.07,0.01,0.1,1.57,0,None,0.12,0),
    ("Feijão preto","Leguminosas",13.6,324,1356,22.9,1.3,0,57.5,8.4,3.3,27,160,1.52,382,8.5,5,1483,0.43,3.56,0,0,0,0.66,0.09,0.41,2.07,3.6,None,0.04,0),
    ("Frango inteiro","Carnes e derivados",72.9,110,460,18.2,3.9,77,0,0,0.9,7,26,0.02,206,0.67,73,284,0.06,1.79,36,36,9,0.06,0.14,0.59,9.54,2.4,None,0.56,0.34),
    ("Carne bovina patinho","Carnes e derivados",72.7,125,523,22.4,4.1,64,0,0,1,9,25,0.01,211,3.02,57,373,0.09,4.36,0,0,0,0.05,0.21,0.35,4.74,0,None,0.4,2.2),
    ("Leite desnatado","Leite e derivados",90.9,35,146,3.4,0.2,0,4.8,0,0.8,125,12,0.01,102,0.04,43,180,0.01,0.43,30,30,9,0.04,0.16,0.04,0.09,1.3,None,0.01,0.45),
    ("Queijo mussarela","Leite e derivados",51.9,264,1105,18.7,19.7,70,2.4,0,3.5,475,21,0.01,324,0.27,396,88,0.01,2.79,191,191,57,0.03,0.33,0.09,0.08,0,None,0.36,1.15),
    ("Pão de forma","Cereais e derivados",35.4,270,1130,8,3.4,0,50.4,2.3,1.6,87,20,0.4,96,1.88,443,113,0.12,0.74,0,0,0,0.2,0.14,0.05,2.45,0,None,0.29,0),
    ("Macarrão cozido","Cereais e derivados",71.7,112,469,4.1,0.5,0,23,1.3,0.2,3,13,0.5,47,0.5,2,35,0.06,0.4,0,0,0,0.02,0.01,0.02,0.54,0,None,0,0),
    ("Soja grão","Leguminosas",6.4,401,1677,36.9,19.9,0,30.7,9.8,4.9,239,220,2.52,570,8.6,1,2008,0.42,3.33,2,2,1,0.9,0.29,0.49,2.03,0,None,1.12,0),
    ("Amido de milho","Cereais e derivados",12,359,1502,0.3,0,0,87.7,0.1,0,1,1,0,11,0.1,0,3,0,0,0,0,0,0,0,0,0,0,None,0,0),
    ("Aveia em flocos","Cereais e derivados",8.1,394,1648,13.9,8.5,0,66.6,9.1,1.9,54,119,3.48,523,4.35,4,371,0.49,3.53,2,2,1,0.89,0.16,0.1,1.26,0,None,0.71,0),
    ("Leite em pó integral","Leite e derivados",2.7,496,2075,26.1,26.3,92,38.4,0,5.8,890,87,0.02,726,0.35,336,1200,0.06,3.12,378,378,112,0.28,1.18,0.34,0.61,6.5,None,0.32,3.2),
    ("Suco de laranja","Frutas",89.2,45,188,0.9,0.3,0,10.2,0.4,0.3,10,12,0.03,20,0.26,0,200,0.06,0.04,22,22,13,0.09,0.03,0.06,0.35,42.5,None,0.15,0),
    ("Chocolate em pó 50%","Outros",3.3,391,1636,13.7,22.9,0,50.9,14.3,5.4,70,299,2.59,419,8.3,15,1081,0.44,4.48,0,0,0,0.09,0.11,0.05,1.71,0,None,0.29,0),
    ("Mel","Açúcares e produtos",17.1,309,1293,0.4,0,0,82.4,0.2,0.2,7,3,0.1,6,0.28,10,164,0.04,0.15,0,0,0,0,0.04,0.07,0.19,0.5,None,None,0),
    ("Farinha de mandioca","Cereais e derivados",7.2,362,1514,1.6,0.5,0,88.4,6.1,0.7,19,40,0.57,40,1.58,6,349,0.1,0.38,0,0,0,0.05,0.06,0.22,1.53,0,None,0.43,0),
    ("Queijo parmesão","Leite e derivados",19.7,451,1887,37,31,89,2.8,0,5.3,1109,40,0.02,777,0.77,1376,116,0.06,4.29,259,259,77,0.05,0.52,0.07,0.31,0,None,0.75,1.2),
    ("Creme de leite","Leite e derivados",62.3,295,1234,2.3,30.4,90,4.1,0,0.6,89,9,0.01,70,0.03,39,119,0.01,0.37,315,315,94,0.03,0.14,0.03,0.07,0.9,None,1.6,0.2),
    ("Iogurte natural","Leite e derivados",87.2,51,213,3.9,1.5,7,5.5,0,0.8,148,14,0.01,111,0.04,64,183,0.01,0.52,23,23,7,0.04,0.2,0.05,0.1,0.9,None,0.06,0.4),
    ("Salmão cozido","Peixes e frutos do mar",63.7,183,766,23.7,10.2,59,0,0,1.2,16,37,0.02,304,0.91,50,414,0.04,0.75,37,37,9,0.24,0.36,0.63,8.01,4.8,16.7,1.82,4.2),
    ("Atum em lata","Peixes e frutos do mar",73.7,95,397,21.5,0.8,43,0,0,1.2,10,31,0.02,285,1.3,296,299,0.04,0.77,0,0,0,0.02,0.06,0.29,11.47,0,None,0.47,2.2),
    ("Espinafre cru","Hortaliças",91.6,22,92,2.7,0.4,0,1.8,2.4,1.5,99,74,0.85,52,2.65,76,554,0.08,0.53,681,469,469,0.09,0.2,0.24,0.72,28.1,None,1.89,0),
    ("Alho","Temperos e molhos",65,132,552,6.4,0.5,0,26.3,4.3,1.5,181,25,1.68,153,1.7,17,600,0.3,1.16,0,0,0,0.2,0.11,1.24,0.7,31.2,None,0.08,0),
    ("Cebola","Hortaliças",90.7,34,142,1.2,0.1,0,7.5,1.8,0.4,25,11,0.17,29,0.2,2,167,0.07,0.2,0,0,0,0.04,0.02,0.12,0.1,7.4,None,0.39,0),
    ("Azeite de oliva","Óleos e gorduras",0,884,3699,0,100,0,0,0,0,0,0,None,0,0,0,0,None,None,0,0,0,0,0,0,0,0,None,14.35,0),
    ("Limão tahiti","Frutas",91.6,33,138,0.9,0.3,0,7.4,2.6,0.4,9,7,0.03,11,0.27,0,130,0.04,0.13,11,11,6,0.04,0.02,0.09,0.24,51.6,None,0.24,0),
    ("Abacate","Frutas",72.3,96,402,1.2,8.4,0,6.4,6.3,0.9,13,29,0.19,52,0.61,4,485,0.17,0.6,10,10,6,0.07,0.14,0.29,1.89,12.5,None,2.07,0),
    ("Brócolis cru","Hortaliças",89.1,34,142,2.8,0.4,0,4.4,2.6,0.8,44,25,0.22,67,0.73,14,325,0.05,0.41,56,56,50,0.07,0.13,0.2,0.6,93.2,None,1.45,0),
    ("Feijão carioca","Leguminosas",14.7,333,1394,22.1,1.4,0,56.8,22.7,3.5,87,190,1.25,408,8.4,5,1637,0.53,3.31,0,0,0,0.62,0.14,0.42,2.01,4.2,None,0.15,0),
    ("Peixe tilápia","Peixes e frutos do mar",78.6,96,402,20,1.7,48,0,0,1.1,14,28,0.02,212,0.57,56,377,0.04,0.66,0,0,0,0.06,0.07,0.13,3.9,2.6,None,0.38,1.1),
    ("Amêndoa sem sal","Oleaginosas",4.7,579,2423,18.7,49.4,0,22.7,11.8,2.8,263,270,2.75,484,3.7,11,732,1.03,3.15,0,0,0,0.18,0.73,0.14,3.87,0,None,26.22,0),
    ("Amendoim torrado","Oleaginosas",2,567,2372,28.1,49,0,16,8.5,2.6,92,168,1.93,376,2.26,9,658,0.67,3.27,0,0,0,0.09,0.11,0.26,13.35,0,None,7.74,0),
    ("Milho verde","Cereais e derivados",73.4,87,364,3.2,1.3,0,20.1,2.3,0.5,3,37,0.16,89,0.5,2,270,0.07,0.69,15,15,9,0.18,0.06,0.12,1.77,12.6,None,0.22,0),
    ("Requeijão cremoso","Leite e derivados",64.6,258,1079,7.9,23.9,83,1.9,0,1.9,186,14,0.01,195,0.18,452,168,0.01,1.1,209,209,63,0.03,0.22,0.04,0.07,0,None,0.47,0.38),
    ("Margarina","Óleos e gorduras",17.6,736,3081,0.5,82.2,0,0.2,0,0.9,11,2,0.01,10,0.06,862,31,0.01,0.09,756,756,0,0,0.01,0,0,0,None,17.17,0),
    ("Ketchup","Molhos e condimentos",75.3,98,410,1.8,0.2,0,21.9,1.1,3.9,22,19,0.32,38,1.51,1035,397,0.12,0.29,101,101,79,0.04,0.09,0.18,0.95,21.7,None,1.55,0),
    ("Vinagre","Temperos e molhos",94.7,18,75,0,0,0,0.2,0,0.1,6,4,0.33,8,0.25,4,73,0.01,0.02,0,0,0,0.01,0.01,0,0,0,None,0.08,0),
    ("Trigo integral farinha","Cereais e derivados",11.3,343,1435,12.6,2,0,71.7,10,2.4,22,107,3.14,296,2.85,2,392,0.36,2.57,0,0,0,0.34,0.12,0.34,3.78,0,None,1.44,0),
]

# ─── Suplemento: alimentos não cobertos pela TACO 4ª edição ──────────────────
# Fontes: TBCA-USP (tbca.fsp.usp.br) e USDA FoodData Central adaptado.
# Formato: (nome, categoria, umidade, kcal, kj, prot, lip, col, carb, fibra, cinzas,
#           ca, mg, mn, p, fe, na, k, cu, zn, retinol, re, rae, tia, ribo, pirid,
#           niac, vitc, vitd, vite, vitb12, fonte)
TACO_SUPLEMENTO = [
    # ─── BASE DE FORMULAÇÃO ───────────────────────────────────────────────────
    # Água de processo: sem nutrientes, mas entra no peso da receita (dilui os
    # valores por 100 g). Sem ela não dá para formular bebidas, caldas e massas.
    ("Água potável","Bebidas",100.0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,"Valor de referência"),
    # ─── LATICÍNIOS ESPECIALIZADOS ────────────────────────────────────────────
    ("Queijo edam","Leite e derivados",44.0,357,1494,24.4,27.8,89,1.4,0,3.8,731,30,0.01,536,0.44,965,188,0.04,3.75,186,186,56,0.03,0.39,0.07,0.06,0,None,0.24,1.54,"USDA-adaptado"),
    ("Queijo gouda","Leite e derivados",41.5,356,1490,24.9,27.4,114,2.2,0,3.9,700,29,0.01,546,0.24,819,121,0.04,3.28,165,165,49,0.03,0.31,0.08,0.07,0,None,0.24,1.54,"USDA-adaptado"),
    ("Queijo brie","Leite e derivados",48.4,334,1397,20.8,27.7,100,0.5,0,2.7,184,20,0.34,188,0.5,629,152,0.02,2.38,592,592,174,0.07,0.52,0.23,0.52,0,None,0.24,1.65,"USDA-adaptado"),
    ("Queijo camembert","Leite e derivados",51.8,300,1255,19.8,24.3,72,0.5,0,3.6,388,20,0.38,347,0.33,842,187,0.02,2.38,820,820,240,0.03,0.49,0.23,0.63,0,None,0.21,1.3,"USDA-adaptado"),
    ("Queijo gorgonzola","Leite e derivados",42.7,353,1477,21.4,28.7,75,2.3,0,5.1,528,23,0.01,311,0.43,1146,256,0.06,2.66,440,440,132,0.03,0.43,0.32,0.49,0,None,0.35,1.22,"USDA-adaptado"),
    ("Queijo emmental","Leite e derivados",37.7,380,1590,28.4,29.0,94,1.6,0,3.4,1080,36,0.01,628,0.19,450,156,0.03,4.34,256,256,77,0.06,0.36,0.05,0.08,0,None,1.72,3.43,"USDA-adaptado"),
    ("Queijo cheddar","Leite e derivados",36.8,402,1682,24.9,33.1,105,1.3,0,3.4,721,27,0.01,512,0.68,621,98,0.03,3.64,994,994,262,0.03,0.43,0.08,0.07,0,None,0.27,1.1,"USDA-adaptado"),
    ("Queijo cottage","Leite e derivados",79.8,98,410,11.1,4.3,17,3.4,0,1.3,83,8,0.01,159,0.07,364,84,0.04,0.47,37,37,11,0.02,0.16,0.07,0.14,0,None,0.08,0.43,"USDA-adaptado"),
    ("Queijo ricota","Leite e derivados",71.7,174,728,11.3,13.0,51,3.0,0,0.9,207,11,0.01,158,0.23,84,105,0.03,1.16,117,117,35,0.01,0.19,0.04,0.09,0,None,0.39,0.34,"USDA-adaptado"),
    ("Queijo provolone","Leite e derivados",40.9,352,1473,25.6,26.6,69,2.1,0,4.8,756,29,0.01,469,0.52,876,138,0.03,3.23,231,231,69,0.03,0.39,0.07,0.06,0,None,0.22,1.54,"USDA-adaptado"),
    ("Queijo prato","Leite e derivados",41.7,364,1524,23.4,29.4,91,1.1,0,3.2,833,27,0.01,500,0.27,1101,93,0.02,3.46,261,261,78,0.03,0.38,0.07,0.07,0,None,0.28,1.15,"TBCA-USP"),
    ("Queijo coalho","Leite e derivados",45.2,267,1117,18.0,20.0,71,2.8,0,3.9,560,22,0.01,380,0.26,704,110,0.02,2.60,200,200,60,0.03,0.30,0.06,0.06,0,None,0.22,0.90,"TBCA-USP"),
    ("Queijo minas frescal","Leite e derivados",58.2,264,1105,17.4,20.2,60,1.6,0,2.0,590,19,0.01,330,0.28,390,98,0.02,2.10,175,175,53,0.03,0.29,0.06,0.07,0,None,0.20,0.85,"TBCA-USP"),
    ("Queijo minas padrão","Leite e derivados",46.5,302,1264,22.0,22.5,75,2.0,0,3.0,710,24,0.01,430,0.30,550,115,0.02,2.80,210,210,63,0.03,0.33,0.06,0.07,0,None,0.24,0.95,"TBCA-USP"),
    ("Queijo tilsit","Leite e derivados",45.4,340,1423,24.4,26.0,102,1.9,0,3.5,700,28,0.01,500,0.40,700,150,0.03,3.50,272,272,82,0.04,0.40,0.08,0.08,0,None,0.26,1.30,"USDA-adaptado"),
    ("Requeijão culinário","Leite e derivados",63.0,240,1004,9.0,21.0,78,2.5,0,2.0,220,14,0.01,200,0.20,500,155,0.01,1.20,200,200,60,0.03,0.20,0.04,0.07,0,None,0.45,0.35,"TBCA-USP"),
    ("Cream cheese","Leite e derivados",53.8,350,1464,5.9,34.4,110,2.7,0,3.2,98,9,0.01,107,0.11,321,138,0.04,0.51,383,383,115,0.02,0.17,0.05,0.07,0,None,0.60,0.22,"USDA-adaptado"),
    ("Manteiga sem sal","Óleos e gorduras",16.2,717,3000,0.85,81.1,215,0.06,0,0.04,24,2,0.00,24,0.02,11,24,0.00,0.09,684,684,671,0.00,0.03,0.00,0.04,0,None,2.32,0.17,"USDA-adaptado"),
    ("Ghee manteiga clarificada","Óleos e gorduras",0.2,900,3766,0.3,99.5,256,0,0,0,4,0,None,3,0.07,2,5,None,0.09,840,840,840,0,0.02,0,0,0,None,2.80,0.3,"USDA-adaptado"),
    ("Creme de leite UHT 35%","Leite e derivados",58.0,345,1443,2.2,36.0,95,3.8,0,0.5,88,9,0.01,69,0.03,35,115,0.01,0.36,385,385,115,0.03,0.14,0.03,0.07,0.5,None,1.6,0.20,"TBCA-USP"),
    ("Creme de leite UHT 20%","Leite e derivados",68.5,196,820,2.5,20.2,72,3.5,0,0.5,90,9,0.01,70,0.03,37,117,0.01,0.37,210,210,63,0.03,0.14,0.04,0.07,0.5,None,0.9,0.20,"TBCA-USP"),
    ("Leite condensado","Leite e derivados",27.2,321,1343,7.9,8.8,34,55.5,0,1.9,290,24,0.01,253,0.19,134,371,0.05,0.97,81,81,24,0.09,0.38,0.12,0.22,1.4,None,0.25,0.54,"TBCA-USP"),
    ("Doce de leite","Leite e derivados",27.0,295,1234,6.7,7.5,28,56.0,0,1.6,250,20,0.01,210,0.18,115,340,0.04,0.85,68,68,20,0.07,0.30,0.10,0.20,1.2,None,0.20,0.45,"TBCA-USP"),
    ("Leite UHT integral","Leite e derivados",88.5,62,259,3.2,3.3,11,4.7,0,0.7,120,12,0.01,92,0.05,42,145,0.01,0.41,38,38,11,0.04,0.18,0.04,0.10,1.0,None,0.06,0.44,"TBCA-USP"),
    ("Leite UHT semidesnatado","Leite e derivados",89.8,48,201,3.3,1.5,6,4.7,0,0.8,122,12,0.01,95,0.05,44,148,0.01,0.42,20,20,6,0.04,0.18,0.04,0.10,0.8,None,0.03,0.45,"TBCA-USP"),
    ("Leite UHT desnatado","Leite e derivados",90.5,36,151,3.4,0.2,2,4.9,0,0.8,126,13,0.01,102,0.05,52,162,0.01,0.43,1,1,0,0.04,0.19,0.04,0.09,0.5,None,0.01,0.45,"TBCA-USP"),
    ("Leite em pó desnatado","Leite e derivados",3.2,354,1481,35.1,0.6,11,52.0,0,8.2,1257,110,0.02,967,0.46,503,1794,0.08,4.50,9,9,3,0.38,1.52,0.46,0.85,7.0,None,0.07,3.50,"TBCA-USP"),
    ("Leite em pó semidesnatado","Leite e derivados",2.9,425,1778,30.0,13.0,50,45.0,0,7.0,1070,100,0.02,840,0.40,420,1490,0.07,3.80,190,190,57,0.33,1.35,0.40,0.73,6.8,None,0.19,3.30,"TBCA-USP"),
    ("Iogurte integral","Leite e derivados",87.0,61,255,3.5,3.3,13,4.7,0,0.7,150,14,0.01,110,0.04,49,190,0.01,0.55,28,28,8,0.04,0.18,0.05,0.10,0.6,None,0.06,0.40,"TBCA-USP"),
    ("Iogurte desnatado","Leite e derivados",88.8,37,155,3.6,0.4,2,4.9,0,0.8,155,14,0.01,114,0.04,52,196,0.01,0.57,5,5,1,0.04,0.19,0.05,0.10,0.4,None,0.01,0.42,"TBCA-USP"),
    ("Iogurte grego integral","Leite e derivados",81.1,97,406,5.7,5.0,17,7.8,0,0.8,125,11,0.01,95,0.05,46,141,0.01,0.52,44,44,13,0.04,0.20,0.04,0.10,0.5,None,0.08,0.75,"TBCA-USP"),
    ("Iogurte grego desnatado","Leite e derivados",82.5,68,284,8.7,0.7,4,5.5,0,0.8,128,12,0.01,98,0.05,47,145,0.01,0.53,10,10,3,0.04,0.20,0.04,0.10,0.4,None,0.02,0.76,"TBCA-USP"),
    ("Iogurte com frutas","Leite e derivados",79.0,90,377,3.3,2.0,9,14.5,0.2,0.7,130,13,0.01,105,0.04,53,175,0.01,0.50,20,20,6,0.04,0.17,0.04,0.09,1.5,None,0.05,0.38,"TBCA-USP"),
    ("Bebida láctea fermentada","Leite e derivados",85.5,67,280,2.8,1.8,8,10.2,0,0.6,105,11,0.01,88,0.04,44,162,0.01,0.42,18,18,5,0.03,0.14,0.04,0.09,1.8,None,0.05,0.30,"TBCA-USP"),
    ("Kefir integral","Leite e derivados",88.0,61,255,3.3,3.5,10,4.5,0,0.7,120,12,0.01,105,0.05,50,150,0.01,0.45,35,35,10,0.04,0.17,0.05,0.10,1.0,None,0.06,0.50,"USDA-adaptado"),
    ("Creme de queijo processado","Leite e derivados",52.0,270,1130,14.0,22.0,68,5.0,0,4.0,550,20,0.01,380,0.28,980,130,0.03,2.50,220,220,66,0.03,0.32,0.06,0.07,0,None,0.30,0.80,"TBCA-USP"),
    ("Queijo processado fatia","Leite e derivados",45.0,292,1222,15.0,21.0,65,10.0,0,6.3,580,18,0.01,400,0.24,1100,120,0.02,2.30,180,180,54,0.03,0.28,0.06,0.06,0,None,0.28,0.75,"TBCA-USP"),
    ("Sobremesa láctea","Leite e derivados",72.0,120,502,3.2,3.5,14,19.5,0,0.7,105,10,0.01,90,0.04,46,140,0.01,0.44,28,28,8,0.03,0.15,0.04,0.08,0.5,None,0.06,0.30,"TBCA-USP"),
    ("Mousse de chocolate lácteo","Leite e derivados",58.0,190,795,4.5,9.5,32,23.0,0.8,1.5,120,18,0.30,110,0.78,75,185,0.07,0.90,58,58,17,0.04,0.22,0.05,0.50,0.4,None,0.30,0.55,"TBCA-USP"),
    # ─── CONCENTRADOS E ISOLADOS PROTEICOS ───────────────────────────────────
    ("Whey protein concentrado 80%","Leite e derivados",4.5,373,1561,80.0,7.0,30,3.5,0,3.5,600,90,0.05,650,0.90,200,580,0.10,4.50,0,0,0,0.30,0.60,0.30,1.20,0,None,0.50,2.50,"USDA-adaptado"),
    ("Whey protein isolado 90%","Leite e derivados",4.0,380,1590,90.0,1.5,15,1.0,0,3.0,650,100,0.05,700,0.80,150,550,0.08,4.80,0,0,0,0.28,0.65,0.28,1.10,0,None,0.45,2.80,"USDA-adaptado"),
    ("Whey protein hidrolisado","Leite e derivados",4.0,378,1582,88.0,2.0,12,2.0,0,3.5,620,95,0.05,680,0.85,160,530,0.08,4.60,0,0,0,0.27,0.62,0.27,1.08,0,None,0.44,2.70,"USDA-adaptado"),
    ("Caseína micelar","Leite e derivados",5.0,360,1507,82.0,2.5,20,4.0,0,5.5,1050,85,0.04,720,0.70,280,520,0.07,4.20,0,0,0,0.20,0.55,0.25,0.90,0,None,0.30,2.40,"USDA-adaptado"),
    ("Caseína hidrolisada","Leite e derivados",5.0,358,1498,81.0,2.0,18,5.0,0,5.5,1000,82,0.04,700,0.70,260,510,0.07,4.10,0,0,0,0.19,0.54,0.24,0.88,0,None,0.29,2.35,"USDA-adaptado"),
    ("Proteína de soro de leite WPC34","Leite e derivados",4.0,360,1507,34.0,4.0,22,52.0,0,5.0,700,80,0.03,550,0.40,250,680,0.07,3.50,0,0,0,0.20,0.40,0.20,0.70,1.5,None,0.28,1.50,"USDA-adaptado"),
    ("Concentrado proteico de leite MPC80","Leite e derivados",4.5,360,1507,80.0,1.5,15,8.0,0,5.5,1200,100,0.04,780,0.60,350,600,0.07,4.80,0,0,0,0.25,0.60,0.28,1.00,0,None,0.38,2.60,"USDA-adaptado"),
    ("Proteína isolada de soja","Leguminosas",5.0,338,1415,88.3,3.4,0,0.7,0.1,3.4,178,31,0.99,776,14.5,1,38,2.51,4.78,0,0,0,0.07,0.14,0.29,1.49,0,None,0.01,0,"USDA-adaptado"),
    ("Proteína concentrada de soja","Leguminosas",6.7,329,1377,65.8,3.3,0,20.7,3.1,4.1,363,32,0.78,744,11.2,5,96,0.65,3.27,0,0,0,0.20,0.24,0.33,3.25,0,None,0.37,0,"USDA-adaptado"),
    ("Proteína texturizada de soja (PTS)","Leguminosas",7.5,327,1368,51.5,1.2,0,34.0,17.5,5.0,241,290,3.09,694,9.46,2,2494,1.97,4.89,0,0,0,0.76,0.13,0.57,2.43,0,None,0.04,0,"USDA-adaptado"),
    ("Colágeno hidrolisado","Carnes e derivados",9.2,376,1573,87.5,0.5,0,0,0,2.5,300,6,0.01,75,0.10,200,15,0.06,0.10,0,0,0,0,0.05,0,0,0,None,0,0,"USDA-adaptado"),
    # ─── LATICÍNIOS ESPECIALIZADOS — GORDURAS E DERIVADOS ────────────────────
    ("Lactose em pó","Leite e derivados",0.1,365,1527,0.3,0.2,0,99.0,0,0.1,120,8,0.01,90,0.04,35,150,0.01,0.30,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Soro de leite em pó","Leite e derivados",3.0,353,1477,11.0,1.0,14,77.9,0,8.0,910,80,0.03,620,0.40,358,1600,0.08,4.50,0,0,0,0.28,1.05,0.35,0.72,6.5,None,0.19,2.40,"USDA-adaptado"),
    ("Permeado de soro (lactossoro)","Leite e derivados",3.2,334,1398,10.0,0.5,6,82.0,0,8.5,950,75,0.03,590,0.35,700,1700,0.07,4.20,0,0,0,0.25,0.90,0.30,0.65,5.0,None,0.15,2.10,"USDA-adaptado"),
    ("Proteína de leite TMP (transmembrana)","Leite e derivados",4.0,372,1557,82.0,3.0,18,5.5,0,5.5,1100,95,0.04,740,0.65,320,580,0.07,4.70,0,0,0,0.24,0.58,0.27,0.98,0,None,0.36,2.55,"USDA-adaptado"),
    ("Butter oil (óleo de manteiga)","Óleos e gorduras",0.5,897,3753,0.2,99.0,256,0.2,0,0,25,2,None,5,0.01,5,10,None,None,840,840,840,0,0.02,0,0,0,None,2.50,0.28,"USDA-adaptado"),
    ("Anidro de gordura do leite (AMF)","Óleos e gorduras",0.2,892,3731,0.1,99.5,255,0.1,0,0,22,1,None,4,0.01,4,8,None,None,824,824,824,0,0.01,0,0,0,None,2.45,0.25,"USDA-adaptado"),
    # ─── CEREAIS E DERIVADOS ESPECIALIZADOS ──────────────────────────────────
    ("Maltodextrina","Cereais e derivados",4.0,380,1590,0.2,0.1,0,95.0,0.5,0.5,3,2,0.01,8,0.05,15,6,0.01,0.03,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Dextrose monoidratada","Cereais e derivados",8.5,364,1524,0,0,0,91.0,0,0.1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Frutose cristalina","Açúcares e produtos",0.3,380,1590,0,0,0,99.6,0,0.1,2,1,0,2,0.05,2,5,0.01,0.01,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Sacarose invertida (xarope 76°Brix)","Açúcares e produtos",23.0,290,1213,0,0,0,76.5,0,0.1,2,1,0,1,0.05,2,5,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Açúcar demerara","Açúcares e produtos",1.5,380,1590,0,0,0,98.0,0,0.5,18,4,0.03,4,0.20,10,80,0.03,0.02,0,0,0,0,0.01,0,0.05,0,None,0,0,"USDA-adaptado"),
    ("Açúcar mascavo","Açúcares e produtos",3.5,356,1490,0.3,0,0,91.8,0.5,1.5,85,29,0.30,22,1.91,28,331,0.30,0.18,0,0,0,0.01,0.01,0.03,0.15,0,None,0,0,"USDA-adaptado"),
    ("Mel de abelha","Açúcares e produtos",17.1,309,1293,0.4,0,0,82.4,0.2,0.2,7,3,0.10,6,0.28,10,164,0.04,0.15,0,0,0,0,0.04,0.07,0.19,0.5,None,None,0,"TACO 4ª edição"),
    ("Glucose de milho (xarope)","Cereais e derivados",16.0,316,1322,0.1,0.1,0,83.5,0,0.3,3,2,0.01,4,0.08,15,12,0.01,0.04,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Amido de mandioca (fécula)","Cereais e derivados",13.5,355,1485,0.5,0.1,0,85.5,1.4,0.4,20,12,0.09,22,0.30,3,310,0.07,0.20,0,0,0,0.02,0.02,0.05,0.40,0,None,0.15,0,"TBCA-USP"),
    ("Amido de milho modificado","Cereais e derivados",11.0,360,1507,0.3,0,0,88.5,0.1,0.2,2,1,0,12,0.12,2,4,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Farinha de arroz","Cereais e derivados",11.9,359,1502,6.0,1.4,0,79.3,2.4,0.5,6,35,1.20,98,0.35,5,76,0.16,1.09,0,0,0,0.23,0.02,0.30,3.04,0,None,0.11,0,"USDA-adaptado"),
    ("Farinha de aveia","Cereais e derivados",8.7,375,1569,13.3,6.9,0,66.6,6.5,1.9,53,138,3.60,452,3.90,6,367,0.46,3.11,0,0,0,0.71,0.10,0.09,0.80,0,None,0.67,0,"USDA-adaptado"),
    ("Farinha de milho","Cereais e derivados",12.0,362,1515,8.1,3.6,0,73.0,7.3,1.7,6,93,0.48,241,2.38,5,315,0.17,1.82,11,11,6,0.36,0.20,0.29,3.63,0,None,0.49,0,"USDA-adaptado"),
    ("Farinha de centeio","Cereais e derivados",10.5,356,1490,8.5,1.6,0,75.9,14.6,1.6,22,72,2.72,279,2.54,1,374,0.29,2.03,0,0,0,0.28,0.07,0.15,2.39,0,None,0.79,0,"USDA-adaptado"),
    ("Farinha de coco","Outras",11.0,440,1841,20.0,18.0,0,26.0,40.0,4.5,26,90,1.00,285,3.70,22,555,0.46,2.50,0,0,0,0.06,0.04,0.08,0.78,1.5,None,0.25,0,"USDA-adaptado"),
    ("Farinha de amêndoa","Oleaginosas",4.4,604,2527,21.4,52.5,0,17.9,12.5,2.4,268,270,2.15,478,3.71,1,733,1.03,3.15,0,0,0,0.20,0.81,0.13,3.89,0,None,26.20,0,"USDA-adaptado"),
    ("Polvilho doce","Cereais e derivados",13.5,342,1431,0.3,0.1,0,85.8,0.8,0.4,15,2,0.05,6,0.10,3,30,0.04,0.10,0,0,0,0.01,0.01,0.02,0.10,0,None,0.05,0,"TBCA-USP"),
    ("Polvilho azedo","Cereais e derivados",12.3,343,1435,0.4,0.1,0,86.2,0.5,0.4,13,2,0.05,5,0.09,3,28,0.04,0.10,0,0,0,0.01,0.01,0.02,0.10,0,None,0.05,0,"TBCA-USP"),
    ("Quinoa grão cru","Cereais e derivados",13.3,368,1540,14.1,6.1,0,64.2,7.0,2.4,47,197,2.03,457,4.57,5,563,0.59,3.10,1,1,0,0.36,0.32,0.49,1.52,0,None,2.44,0,"USDA-adaptado"),
    ("Aveia em flocos finos","Cereais e derivados",8.2,389,1628,16.9,6.9,0,66.3,10.6,1.7,54,177,4.92,523,4.72,2,429,0.63,3.97,0,0,0,0.76,0.14,0.12,0.96,0,None,0.42,0,"USDA-adaptado"),
    ("Granola tradicional","Cereais e derivados",4.2,471,1971,10.8,19.9,0,60.8,5.8,2.1,54,95,2.77,307,3.48,27,295,0.41,2.94,0,0,0,0.41,0.13,0.13,2.68,0.4,None,5.32,0,"USDA-adaptado"),
    ("Gérmen de trigo","Cereais e derivados",13.5,360,1507,23.2,9.7,0,51.8,13.2,4.2,39,239,13.30,842,6.26,12,892,0.80,12.29,0,0,0,1.88,0.49,1.30,6.81,0,None,17.00,0,"USDA-adaptado"),
    # ─── FRUTAS ADICIONAIS ────────────────────────────────────────────────────
    ("Morango","Frutas",91.6,34,142,0.7,0.3,0,7.4,1.8,0.5,16,13,0.39,24,0.41,1,153,0.05,0.14,0,0,0,0.02,0.02,0.05,0.39,58.8,None,0.29,0,"TACO 4ª edição"),
    ("Uva comum","Frutas",81.3,67,280,0.7,0.3,0,17.0,0.9,0.5,8,6,0.07,14,0.25,1,188,0.13,0.08,2,2,1,0.07,0.06,0.08,0.27,3.1,None,0.19,0,"TACO 4ª edição"),
    ("Manga palmer","Frutas",80.0,64,268,0.9,0.2,0,17.1,1.7,0.4,9,9,0.06,12,0.16,2,156,0.11,0.08,389,389,234,0.08,0.05,0.14,0.66,36.4,None,1.12,0,"TACO 4ª edição"),
    ("Mamão formosa","Frutas",90.1,36,151,1.0,0.2,0,8.0,1.8,0.5,23,10,0.04,10,0.07,5,257,0.05,0.08,87,87,52,0.03,0.04,0.02,0.35,74.3,None,0.33,0,"TACO 4ª edição"),
    ("Melancia","Frutas",92.4,33,138,0.7,0.2,0,7.8,0.4,0.3,8,10,0.04,11,0.24,1,112,0.04,0.10,28,28,17,0.03,0.02,0.05,0.18,8.1,None,0.05,0,"TACO 4ª edição"),
    ("Melão","Frutas",90.6,29,121,0.8,0.1,0,7.6,0.3,0.6,8,10,0.04,13,0.19,9,224,0.04,0.16,40,40,24,0.04,0.02,0.08,0.50,25.0,None,0.05,0,"TACO 4ª edição"),
    ("Abacaxi","Frutas",86.2,48,201,0.9,0.1,0,12.3,1.0,0.3,18,18,0.92,8,0.26,1,150,0.08,0.09,2,2,1,0.08,0.04,0.11,0.49,47.8,None,0.02,0,"TACO 4ª edição"),
    ("Goiaba","Frutas",81.0,54,226,2.6,0.5,0,13.2,6.0,0.8,18,10,0.15,25,0.26,2,198,0.23,0.23,57,57,42,0.04,0.03,0.11,1.16,228.3,None,0.73,0,"TACO 4ª edição"),
    ("Caju","Frutas",84.2,43,180,1.3,0.2,0,10.1,1.7,0.5,6,22,0.17,45,0.48,4,208,0.31,0.32,3,3,2,0.02,0.04,0.20,0.63,219.7,None,0.38,0,"TACO 4ª edição"),
    ("Kiwi","Frutas",83.9,61,255,0.9,0.6,0,14.7,3.0,0.7,34,17,0.10,34,0.24,4,312,0.13,0.14,4,4,2,0.03,0.04,0.07,0.35,92.7,None,1.46,0,"USDA-adaptado"),
    ("Pêssego","Frutas",88.9,37,155,0.9,0.1,0,9.5,1.5,0.4,6,9,0.07,20,0.25,0,190,0.07,0.17,16,16,10,0.02,0.03,0.02,0.81,6.6,None,0.73,0,"TACO 4ª edição"),
    ("Pêra","Frutas",84.9,45,188,0.7,0.1,0,13.6,3.1,0.4,11,8,0.05,11,0.20,1,119,0.08,0.10,1,1,0,0.01,0.02,0.02,0.16,3.8,None,0.12,0,"USDA-adaptado"),
    ("Ameixa fresca","Frutas",87.2,46,192,0.7,0.3,0,11.4,1.4,0.5,6,7,0.05,16,0.17,0,157,0.06,0.10,17,17,10,0.03,0.03,0.03,0.42,9.5,None,0.26,0,"USDA-adaptado"),
    ("Figo fresco","Frutas",79.1,74,310,0.8,0.3,0,19.2,2.9,0.9,35,17,0.13,14,0.37,1,232,0.07,0.15,7,7,4,0.06,0.05,0.11,0.40,2.0,None,0.11,0,"USDA-adaptado"),
    ("Coco fresco ralado","Frutas",46.9,354,1482,3.4,34.0,0,14.3,8.9,1.0,14,32,1.50,113,2.43,20,356,0.43,1.10,0,0,0,0.07,0.02,0.05,0.54,3.3,None,0.15,0,"TBCA-USP"),
    ("Acerola","Frutas",92.9,33,138,0.8,0.2,0,7.1,1.5,0.5,12,18,0.09,17,0.22,7,146,0.18,0.10,38,38,23,0.02,0.06,0.03,0.40,941.4,None,0.10,0,"TACO 4ª edição"),
    ("Pitanga","Frutas",90.8,39,163,0.9,0.6,0,8.2,1.3,0.5,18,11,0.16,11,0.26,5,113,0.04,0.19,489,489,254,0.03,0.03,0.03,0.31,26.3,None,1.02,0,"TACO 4ª edição"),
    ("Cupuaçu polpa","Frutas",84.8,51,213,1.5,0.5,0,12.4,2.6,0.6,10,30,0.27,20,1.20,8,200,0.35,0.36,2,2,1,0.04,0.20,0.09,0.96,24.6,None,0.26,0,"TBCA-USP"),
    ("Açaí polpa congelada","Frutas",67.6,247,1034,2.2,18.2,0,7.1,2.4,0.5,81,57,2.50,58,3.57,8,243,1.10,1.42,40,40,24,0.25,0.32,0.15,0.65,4.0,None,4.47,0,"TBCA-USP"),
    ("Tamarindo polpa","Frutas",31.4,239,1000,2.8,0.6,0,62.5,5.1,2.7,74,92,0.93,113,2.80,28,628,0.09,0.12,2,2,1,0.43,0.15,0.07,1.94,3.5,None,0.14,0,"USDA-adaptado"),
    # ─── HORTALIÇAS ADICIONAIS ────────────────────────────────────────────────
    ("Couve manteiga","Hortaliças",90.5,27,113,2.0,0.5,0,4.1,2.0,1.4,150,19,0.66,28,0.90,5,228,0.19,0.39,310,310,267,0.07,0.13,0.27,1.00,76.0,None,2.26,0,"TACO 4ª edição"),
    ("Couve flor","Hortaliças",91.4,25,105,1.9,0.2,0,4.9,2.4,0.7,22,15,0.16,44,0.42,30,303,0.04,0.28,0,0,0,0.05,0.06,0.20,0.51,59.2,None,0.08,0,"TACO 4ª edição"),
    ("Repolho","Hortaliças",92.1,25,105,1.3,0.1,0,5.4,1.8,0.6,43,12,0.16,26,0.47,10,170,0.02,0.15,5,5,4,0.06,0.04,0.12,0.21,47.5,None,0.15,0,"TACO 4ª edição"),
    ("Alface","Hortaliças",96.0,11,46,1.3,0.2,0,1.7,1.7,0.8,36,11,0.39,29,0.86,10,194,0.05,0.15,188,188,163,0.06,0.08,0.08,0.35,7.9,None,0.44,0,"TACO 4ª edição"),
    ("Rúcula","Hortaliças",91.7,25,105,2.6,0.7,0,3.6,1.6,1.6,160,47,0.32,52,1.46,27,369,0.08,0.47,119,119,109,0.04,0.09,0.07,0.31,15.0,None,0.43,0,"USDA-adaptado"),
    ("Pepino","Hortaliças",95.2,11,46,0.6,0.1,0,2.2,0.6,0.5,14,12,0.07,21,0.26,2,136,0.05,0.20,11,11,8,0.03,0.03,0.04,0.10,3.2,None,0.03,0,"TACO 4ª edição"),
    ("Abobrinha","Hortaliças",94.7,18,75,1.4,0.2,0,2.9,1.2,0.5,19,18,0.18,38,0.37,3,262,0.04,0.32,10,10,8,0.04,0.09,0.09,0.45,13.9,None,0.12,0,"TACO 4ª edição"),
    ("Berinjela","Hortaliças",92.0,24,100,1.2,0.2,0,5.5,3.4,0.6,9,14,0.25,25,0.23,2,229,0.08,0.16,1,1,1,0.03,0.03,0.08,0.65,2.2,None,0.30,0,"TACO 4ª edição"),
    ("Pimentão vermelho","Hortaliças",92.2,28,117,1.0,0.3,0,6.0,2.1,0.5,9,12,0.18,28,0.43,3,212,0.07,0.25,157,157,101,0.05,0.09,0.29,0.98,200.0,None,1.58,0,"USDA-adaptado"),
    ("Pimentão verde","Hortaliças",93.9,20,84,0.9,0.2,0,4.3,1.8,0.5,8,10,0.12,24,0.34,2,175,0.06,0.17,18,18,15,0.06,0.04,0.22,0.48,128.0,None,0.69,0,"USDA-adaptado"),
    ("Abóbora moranga","Hortaliças",93.0,24,100,1.1,0.1,0,5.3,0.8,0.5,16,9,0.12,37,0.27,2,382,0.10,0.22,1155,1155,700,0.05,0.11,0.06,0.60,9.0,None,1.06,0,"TBCA-USP"),
    ("Vagem","Hortaliças",91.4,28,117,1.8,0.1,0,6.0,2.7,0.6,38,26,0.22,44,1.03,2,209,0.07,0.24,24,24,20,0.08,0.10,0.07,0.73,16.3,None,0.41,0,"TACO 4ª edição"),
    ("Ervilha fresca","Hortaliças",78.7,72,301,5.4,0.4,0,13.6,5.7,0.9,25,33,0.41,108,1.47,5,244,0.17,1.24,38,38,25,0.27,0.13,0.16,2.09,40.0,None,0.13,0,"TACO 4ª edição"),
    ("Milho cozido","Cereais e derivados",72.7,91,381,3.4,1.4,0,20.7,2.2,0.6,4,34,0.17,92,0.48,3,243,0.07,0.68,11,11,7,0.17,0.05,0.12,1.64,7.6,None,0.21,0,"TACO 4ª edição"),
    ("Mandioca cozida","Cereais e derivados",61.0,125,523,1.0,0.3,0,30.1,1.9,0.7,17,21,0.25,28,0.94,3,271,0.06,0.43,0,0,0,0.06,0.03,0.09,0.65,20.6,None,0.19,0,"TACO 4ª edição"),
    ("Beterraba","Hortaliças",90.5,35,146,1.5,0.1,0,7.6,2.5,0.9,14,23,0.33,38,0.80,58,325,0.07,0.35,0,0,0,0.03,0.04,0.06,0.33,3.1,None,0.04,0,"TACO 4ª edição"),
    ("Jiló","Hortaliças",93.7,19,79,1.0,0.3,0,4.5,1.6,0.6,20,18,0.15,25,0.40,5,180,0.10,0.18,22,22,14,0.04,0.04,0.05,0.55,23.3,None,0.25,0,"TBCA-USP"),
    ("Quiabo","Hortaliças",90.0,33,138,2.0,0.2,0,7.0,3.2,0.9,77,57,0.79,61,0.62,7,303,0.09,0.60,36,36,22,0.20,0.06,0.22,1.00,23.0,None,0.27,0,"TACO 4ª edição"),
    # ─── CARNES E PESCADOS ADICIONAIS ─────────────────────────────────────────
    ("Carne suína lombo","Carnes e derivados",71.2,131,548,22.4,4.6,65,0,0,1.1,14,28,0.01,237,0.95,54,380,0.09,2.93,0,0,0,0.59,0.25,0.40,7.01,0,None,0.36,0.69,"TACO 4ª edição"),
    ("Carne suína costelinha","Carnes e derivados",65.8,167,699,18.5,10.0,72,0,0,1.1,14,25,0.01,213,1.36,62,291,0.07,2.55,0,0,0,0.41,0.22,0.32,5.02,0,None,0.26,0.69,"TACO 4ª edição"),
    ("Carne bovina alcatra","Carnes e derivados",73.5,114,477,22.6,2.9,62,0,0,1.0,5,26,0.01,220,2.78,52,358,0.08,4.10,0,0,0,0.07,0.20,0.34,4.68,0,None,0.38,2.20,"TACO 4ª edição"),
    ("Carne bovina filé mignon","Carnes e derivados",75.0,99,414,22.2,1.5,59,0,0,0.9,4,25,0.01,215,2.72,50,345,0.08,3.98,0,0,0,0.07,0.19,0.33,4.55,0,None,0.37,2.15,"TBCA-USP"),
    ("Carne bovina contrafilé","Carnes e derivados",71.8,127,531,20.8,4.8,65,0,0,1.0,6,24,0.01,195,2.90,52,310,0.07,3.80,0,0,0,0.06,0.18,0.30,4.40,0,None,0.34,2.10,"TBCA-USP"),
    ("Frango peito sem pele","Carnes e derivados",74.2,109,456,23.3,1.6,68,0,0,1.1,5,28,0.02,223,0.64,63,326,0.04,0.80,0,0,0,0.07,0.13,0.64,10.60,0,None,0.27,0.32,"TBCA-USP"),
    ("Frango sobrecoxa sem pele","Carnes e derivados",71.8,125,523,21.0,4.3,77,0,0,1.2,8,26,0.02,200,0.80,78,310,0.07,2.10,0,0,0,0.06,0.16,0.55,6.40,0,None,0.35,0.40,"TBCA-USP"),
    ("Camarão cozido","Peixes e frutos do mar",76.5,99,414,21.4,1.1,152,0,0,1.7,70,34,0.04,237,2.41,224,220,0.26,1.93,0,0,0,0.03,0.03,0.10,2.20,1.9,None,1.20,1.32,"USDA-adaptado"),
    ("Tilápia grelhada","Peixes e frutos do mar",67.0,128,536,26.2,2.7,57,0,0,1.3,10,27,0.02,204,0.56,52,380,0.04,0.63,0,0,0,0.06,0.07,0.16,5.50,0,3.82,0.37,1.20,"USDA-adaptado"),
    ("Atum fresco","Peixes e frutos do mar",70.5,109,456,24.4,0.5,38,0,0,1.3,8,30,0.02,254,1.02,39,407,0.08,0.60,0,0,0,0.24,0.25,0.45,8.65,0,5.40,0.64,9.43,"USDA-adaptado"),
    ("Sardinha em lata ao natural","Peixes e frutos do mar",60.7,208,871,24.6,11.5,142,0,0,2.9,382,35,0.10,490,2.92,505,397,0.23,1.32,54,54,13,0.05,0.22,0.18,5.24,0,7.6,2.04,8.94,"USDA-adaptado"),
    ("Ovo de codorna inteiro","Ovos",74.6,158,661,13.1,11.1,844,0.4,0,1.1,64,13,0.04,226,3.65,141,132,0.11,1.47,156,156,156,0.13,0.79,0.15,0.20,0,1.4,1.08,1.58,"USDA-adaptado"),
    ("Clara de ovo","Ovos",87.6,52,218,10.9,0.2,0,0.7,0,0.6,7,11,0.01,15,0.08,166,163,0.02,0.03,0,0,0,0.02,0.45,0.01,0.10,0,None,0,0.09,"USDA-adaptado"),
    ("Gema de ovo","Ovos",52.3,322,1347,15.9,26.5,1085,0.6,0,1.8,129,16,0.08,390,2.73,48,109,0.06,2.30,537,537,537,0.18,0.53,0.35,0.02,0,2.18,2.58,3.79,"USDA-adaptado"),
    # ─── LEGUMINOSAS ADICIONAIS ───────────────────────────────────────────────
    ("Lentilha cozida","Leguminosas",69.6,114,477,9.0,0.4,0,20.1,7.9,0.8,19,36,0.49,180,3.33,2,369,0.25,1.27,1,1,1,0.17,0.07,0.18,1.06,1.5,None,0.11,0,"USDA-adaptado"),
    ("Grão-de-bico cozido","Leguminosas",60.2,164,686,8.9,2.6,0,27.4,7.6,1.4,49,48,1.03,168,2.89,7,291,0.35,1.53,1,1,1,0.12,0.06,0.14,0.53,1.3,None,0.35,0,"USDA-adaptado"),
    ("Feijão fradinho","Leguminosas",11.0,335,1402,23.3,1.7,0,58.7,6.3,3.4,54,214,1.02,416,9.03,6,1345,0.84,3.27,0,0,0,0.73,0.23,0.25,1.82,0.5,None,0.23,0,"TBCA-USP"),
    ("Ervilha seca","Leguminosas",11.0,338,1414,23.8,1.2,0,60.4,14.6,2.8,55,115,0.81,366,4.43,15,825,0.87,3.30,6,6,4,0.73,0.22,0.14,3.05,40.0,None,0.08,0,"USDA-adaptado"),
    # ─── OLEAGINOSAS ADICIONAIS ───────────────────────────────────────────────
    ("Castanha-do-Pará","Oleaginosas",3.5,656,2745,14.3,66.4,0,12.3,7.5,3.5,160,376,0.81,725,2.43,3,659,1.74,4.06,0,0,0,0.62,0.04,0.28,0.30,0.7,None,5.73,0,"TACO 4ª edição"),
    ("Castanha de caju","Oleaginosas",5.2,570,2385,18.5,46.4,0,28.7,3.3,2.7,37,292,1.66,593,6.68,12,660,2.22,5.78,0,0,0,0.42,0.06,0.42,1.40,0.5,None,0.90,0,"TACO 4ª edição"),
    ("Nozes","Oleaginosas",4.1,620,2595,15.2,59.4,0,18.3,6.7,1.8,98,158,3.41,346,2.91,2,441,1.59,3.09,0,0,0,0.34,0.15,0.54,1.12,1.3,None,0.70,0,"USDA-adaptado"),
    ("Pistache sem casca","Oleaginosas",4.0,562,2352,20.2,45.3,0,27.5,10.3,2.8,105,121,1.24,490,3.92,1,1025,1.30,2.20,0,0,0,0.87,0.16,1.70,1.30,5.6,None,2.86,0,"USDA-adaptado"),
    ("Macadâmia","Oleaginosas",1.4,718,3004,7.9,75.8,0,13.8,8.6,1.2,85,130,4.13,188,3.69,5,368,0.76,1.30,0,0,0,1.20,0.16,0.28,2.47,1.2,None,0.54,0,"USDA-adaptado"),
    ("Tahine (pasta de gergelim)","Oleaginosas",3.0,595,2490,17.0,53.8,0,21.2,9.3,3.5,426,95,1.05,573,8.95,115,414,1.49,4.62,3,3,2,0.24,0.47,0.15,5.45,0,None,0.25,0,"USDA-adaptado"),
    ("Linhaça dourada","Oleaginosas",6.9,534,2235,18.3,42.2,0,28.8,27.3,3.4,255,392,2.48,642,5.73,30,813,1.22,4.34,0,0,0,1.64,0.16,0.47,3.08,0.6,None,31.51,0,"USDA-adaptado"),
    ("Chia semente","Oleaginosas",5.8,486,2034,16.5,30.7,0,42.1,34.4,4.7,631,335,2.72,860,7.72,16,407,0.92,4.58,0,0,0,0.62,0.17,0.19,8.83,1.6,None,0.50,0,"USDA-adaptado"),
    ("Gergelim","Oleaginosas",4.7,573,2397,17.7,49.7,0,23.4,11.8,4.9,975,351,2.46,629,14.55,11,468,4.08,7.75,0,0,0,0.79,0.25,0.79,4.52,0,None,0.25,0,"USDA-adaptado"),
    # ─── ÓLEOS ESPECIALIZADOS ─────────────────────────────────────────────────
    ("Óleo de coco","Óleos e gorduras",0,892,3731,0,100,0,0,0,0,1,0,None,0,0,0,0,None,None,0,0,0,0,0,0,0,0,None,0.11,0,"USDA-adaptado"),
    ("Óleo de girassol","Óleos e gorduras",0,884,3699,0,100,0,0,0,0,0,0,None,0,0,0,0,None,None,0,0,0,0,0,0,0,0,None,41.08,0,"USDA-adaptado"),
    ("Óleo de canola","Óleos e gorduras",0,884,3699,0,100,0,0,0,0,0,0,None,0,0,0,0,None,None,0,0,0,0,0,0,0,0,None,17.46,0,"USDA-adaptado"),
    ("Óleo de milho","Óleos e gorduras",0,884,3699,0,100,0,0,0,0,0,0,None,0,0,0,0,None,None,0,0,0,0,0,0,0,0,None,14.30,0,"USDA-adaptado"),
    ("Óleo de palma (dendê)","Óleos e gorduras",0,884,3699,0,100,0,0,0,0,0,0,None,0,0,0,0,None,None,500,500,500,0,0,0,0,0,None,15.94,0,"USDA-adaptado"),
    ("Óleo de oliva extra virgem","Óleos e gorduras",0,884,3699,0,100,0,0,0,0,0,0,None,0,0,0,0,None,None,0,0,0,0,0,0,0,0,None,14.35,0,"USDA-adaptado"),
    ("Óleo de linhaça","Óleos e gorduras",0,884,3699,0,100,0,0,0,0,0,0,None,0,0,0,0,None,None,0,0,0,0,0,0,0,0,None,17.50,0,"USDA-adaptado"),
    ("Gordura vegetal hidrogenada","Óleos e gorduras",0,884,3699,0,100,0,0,0,0,0,0,None,0,0,0,0,None,None,0,0,0,0,0,0,0,0,None,12.00,0,"USDA-adaptado"),
    ("Gordura interesterificada","Óleos e gorduras",0,884,3699,0,100,0,0,0,0,0,0,None,0,0,0,0,None,None,0,0,0,0,0,0,0,0,None,10.00,0,"USDA-adaptado"),
    # ─── ADITIVOS E INGREDIENTES INDUSTRIAIS ─────────────────────────────────
    ("Lecitina de soja","Aditivos e ingredientes",2.0,763,3193,2.9,96.1,0,0,0,0,100,40,0.05,480,4.50,28,50,0.18,3.11,0,0,0,0.01,0.05,0.01,0.10,0,None,100.0,0,"USDA-adaptado"),
    ("Goma xantana","Aditivos e ingredientes",10.0,310,1297,6.0,0.8,0,75.0,74.0,4.6,100,50,0.50,200,0.50,1000,200,0.10,0.50,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Goma guar","Aditivos e ingredientes",10.0,315,1318,5.5,0.5,0,80.0,80.0,3.0,210,170,1.20,260,2.10,40,700,0.38,1.03,0,0,0,0.01,0.03,0.01,0.30,0,None,0.10,0,"USDA-adaptado"),
    ("Carragena","Aditivos e ingredientes",10.0,253,1059,0.5,0.1,0,62.6,62.6,25.5,200,150,0.80,130,5.00,6200,110,0.80,0.50,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Ágar-ágar","Aditivos e ingredientes",16.0,306,1280,6.2,0.3,0,81.0,81.0,0,625,770,4.31,52,21.40,102,1125,0.28,5.77,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Pectina","Aditivos e ingredientes",7.6,314,1314,3.1,0,0,87.4,87.4,0,160,50,0.30,60,3.00,60,290,0.80,0.60,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Gelatina incolor em pó","Carnes e derivados",13.0,335,1402,85.6,0.1,0,0,0,1.3,50,10,0.01,60,0.30,130,10,0.05,0.09,0,0,0,0,0.08,0,0,0,None,0,0,"USDA-adaptado"),
    ("Carboximetilcelulose (CMC)","Aditivos e ingredientes",5.0,300,1255,0,0,0,99.0,99.0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Metilcelulose","Aditivos e ingredientes",5.0,296,1239,0,0,0,99.0,99.0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Hidroxipropilmetilcelulose (HPMC)","Aditivos e ingredientes",5.0,280,1172,0,0,0,93.0,93.0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Amido resistente (RS tipo II)","Cereais e derivados",10.0,360,1507,0.5,0.1,0,89.0,89.0,0,10,5,0.02,20,0.20,5,15,0.01,0.05,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Inulina","Aditivos e ingredientes",3.0,200,837,0.2,0.1,0,90.0,90.0,0,35,15,0.05,30,0.20,5,95,0.05,0.10,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("FOS (fruto-oligossacarídeos)","Aditivos e ingredientes",3.0,150,628,0.1,0,0,80.0,80.0,0,20,8,0.03,15,0.15,3,50,0.03,0.05,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("GOS (galacto-oligossacarídeos)","Aditivos e ingredientes",4.0,175,733,0.1,0,0,88.0,88.0,0,50,10,0.02,20,0.05,10,40,0.02,0.04,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Sorbitol","Aditivos e ingredientes",0.5,254,1063,0,0,0,99.0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Sucralose","Aditivos e ingredientes",0,0,0,0,0,0,100.0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Estévia (extrato 97% rebaudiosdeo A)","Aditivos e ingredientes",4.0,284,1188,0,0,0,95.0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Aspartame","Aditivos e ingredientes",1.0,352,1473,57.1,0,0,42.9,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Ciclamato de sódio","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Sacarina sódica","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Acesulfame-K","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Ácido cítrico anidro","Aditivos e ingredientes",0,247,1033,0,0,0,73.0,0,0,13,2,0.05,14,0.17,3,43,0.05,0.05,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Ácido láctico","Aditivos e ingredientes",28.0,260,1088,0,0,0,72.0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Ácido acético glacial","Aditivos e ingredientes",0,209,875,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Ácido málico","Aditivos e ingredientes",0,234,979,0,0,0,100.0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Ácido tartárico","Aditivos e ingredientes",0,249,1042,0,0,0,100.0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Bicarbonato de sódio","Aditivos e ingredientes",0,0,0,0,0,0,0,0,65.0,0,0,0,0,0,27000,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Cloreto de sódio (sal puro)","Condimentos",0,0,0,0,0,0,0,0,60.3,0,0,0,0,0,38758,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Cloreto de potássio (KCl)","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,52000,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Fosfato tricálcico","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,2990,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Citrato de cálcio","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,2100,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Sulfato ferroso","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,200000,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Gluconato de zinco","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,142857,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Vitamina C (ácido ascórbico)","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1000000,None,0,0,"USDA-adaptado"),
    ("Vitamina E (tocoferol)","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,1000000,0,"USDA-adaptado"),
    ("Vitamina D3 (colecalciferol)","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1000000,0,0,"USDA-adaptado"),
    ("Vitamina B12 (cianocobalamina)","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,1000000,"USDA-adaptado"),
    ("Cloreto de colina","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Monoglicerídeo (emulsificante 471)","Aditivos e ingredientes",0,794,3322,0,88.0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,1.0,0,"USDA-adaptado"),
    ("Diglicerídeo (emulsificante 472)","Aditivos e ingredientes",0,780,3264,0,86.0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,1.0,0,"USDA-adaptado"),
    ("Éster de ácido diacetil tartárico (DATEM)","Aditivos e ingredientes",0,750,3138,0,80.0,0,12.0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Éster de sorbitana (Span 60)","Aditivos e ingredientes",0,800,3348,0,85.0,0,15.0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Polissorbato 80","Aditivos e ingredientes",0,505,2113,0,45.0,0,55.0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Dipropionato de sódio (E281)","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Nisina (E234)","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Natamicina (E235)","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Sorbato de potássio (E202)","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Benzoato de sódio (E211)","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Cúrcuma (corante E100)","Especiarias e condimentos",11.4,312,1306,9.7,3.3,0,67.1,22.7,6.6,183,193,19.80,268,41.42,38,2525,1.30,4.35,0,0,0,0.15,0.23,0.19,5.14,25.9,None,3.10,0,"USDA-adaptado"),
    ("Carmim cochinilha (E120)","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Dióxido de titânio (E171)","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Extrato de baunilha","Especiarias e condimentos",34.3,288,1205,0.1,0.1,0,12.7,0,0.3,11,12,0.23,10,0.08,11,148,0.07,0.11,0,0,0,0.01,0.01,0.02,0.40,0,None,0,0,"USDA-adaptado"),
    ("Baunilha em pó","Especiarias e condimentos",30.0,288,1205,0.1,0.1,0,12.7,0,0.3,11,12,0.23,10,0.08,11,148,0.07,0.11,0,0,0,0.01,0.01,0.02,0.40,0,None,0,0,"USDA-adaptado"),
    ("Cacau em pó 100%","Outros",4.4,228,954,19.6,13.7,0,57.9,37.0,5.8,128,499,3.84,734,13.86,21,1524,3.79,6.81,0,0,0,0.07,0.24,0.12,2.19,0,None,0.59,0,"USDA-adaptado"),
    ("Chocolate amargo 70%","Outros",1.5,598,2503,7.8,42.6,0,45.9,10.9,2.3,73,228,1.95,308,11.90,20,715,1.77,3.31,0,0,0,0.05,0.09,0.07,1.05,0,None,1.80,0,"USDA-adaptado"),
    ("Chocolate ao leite","Outros",1.5,535,2239,7.6,29.7,23,59.4,3.4,1.7,214,58,0.42,209,2.35,79,439,0.49,2.28,68,68,20,0.07,0.24,0.05,0.53,0.3,None,1.36,0.37,"USDA-adaptado"),
    ("Chocolate branco","Outros",1.3,539,2256,5.9,32.1,21,59.2,0,2.0,199,24,0.01,173,0.24,90,286,0.04,1.09,193,193,58,0.05,0.30,0.04,0.21,0.5,None,0.99,0.34,"USDA-adaptado"),
    ("Manteiga de amendoim","Oleaginosas",1.1,593,2482,22.2,51.1,0,21.6,6.0,2.8,60,179,1.97,365,1.79,18,649,0.41,2.76,0,0,0,0.12,0.14,0.44,14.08,0,None,7.24,0,"USDA-adaptado"),
    # ─── BEBIDAS ──────────────────────────────────────────────────────────────
    ("Café expresso","Bebidas",99.0,9,38,0.1,0.2,0,1.7,0,0.1,3,6,0.02,14,0.01,2,115,0.01,0.05,0,0,0,0.01,0.18,0,0.26,0,None,0.22,0,"USDA-adaptado"),
    ("Chá preto infusão","Bebidas",99.6,1,4,0.2,0,0,0.3,0,0,0,3,0.52,1,0.02,3,37,0.01,0.02,0,0,0,0.01,0.05,0,0,0,None,0,0,"USDA-adaptado"),
    ("Refrigerante cola diet","Bebidas",99.5,0,1,0.1,0,0,0.2,0,0.1,3,2,0.02,6,0.01,12,4,0.01,0.04,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Água de coco","Bebidas",94.7,19,79,0.7,0.2,0,3.7,1.1,0.6,24,25,0.14,20,0.29,105,250,0.10,0.10,0,0,0,0.03,0.06,0.03,0.11,2.4,None,0.04,0,"TBCA-USP"),
    ("Suco de uva integral","Bebidas",84.4,68,285,0.7,0.1,0,15.6,0.2,0.5,12,8,0.28,14,0.26,5,189,0.04,0.06,1,1,0,0.04,0.05,0.10,0.20,1.0,None,0,0,"USDA-adaptado"),
    ("Leite de amêndoa","Bebidas",98.0,16,67,0.5,1.1,0,0.6,0.2,0.1,188,6,0.04,25,0.35,65,40,0.07,0.06,46,46,14,0,0.04,0,0,0,0.88,6.33,0,"USDA-adaptado"),
    ("Leite de aveia","Bebidas",91.0,48,201,1.5,1.5,0,6.2,0.5,0.5,185,14,0.27,105,0.36,68,107,0.05,0.35,61,61,18,0.08,0.20,0.04,0.10,0,1.06,1.37,0,"USDA-adaptado"),
    ("Bebida vegetal de soja","Bebidas",88.0,43,180,3.4,1.8,0,4.0,0.4,0.6,123,25,0.35,52,0.60,51,118,0.15,0.28,46,46,14,0.07,0.17,0.05,0.20,0.3,1.18,0.18,0.80,"USDA-adaptado"),
    # ─── INGREDIENTES ESPECIALIZADOS P&D LÁCTEOS ──────────────────────────────
    ("Lactase (beta-galactosidase)","Aditivos e ingredientes",5.0,80,335,15.0,0.5,0,10.0,0,0,20,10,0.05,50,0.50,100,50,0.10,0.30,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Coalho bovino líquido","Aditivos e ingredientes",80.0,30,126,5.0,0.1,0,0,0,2.0,5,2,0.01,10,0.10,500,30,0.01,0.10,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Cultura láctica termofílica","Aditivos e ingredientes",5.0,50,209,30.0,2.0,0,5.0,0,5.0,100,20,0.05,80,0.30,200,100,0.05,0.30,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Cultura láctica mesofílica","Aditivos e ingredientes",5.0,50,209,30.0,2.0,0,5.0,0,5.0,100,20,0.05,80,0.30,200,100,0.05,0.30,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Cloreto de cálcio alimentício","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,36220,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Nitrato de sódio (cura)","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Nitrito de sódio (cura)","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Fermento biológico seco","Cereais e derivados",8.0,325,1360,40.4,7.6,0,41.2,26.9,7.5,30,54,0.31,637,3.93,51,955,0.43,7.11,0,0,0,0.33,4.01,0.50,14.01,0,None,0.50,0.04,"USDA-adaptado"),
    ("Fermento biológico fresco","Cereais e derivados",69.0,89,373,12.1,1.9,0,14.6,8.1,1.8,8,16,0.10,194,1.20,15,290,0.14,2.20,0,0,0,0.10,1.22,0.15,4.20,0,None,0.15,0.01,"USDA-adaptado"),
    ("Bicarbonato de amônio","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Caseína ácida","Leite e derivados",12.0,357,1494,84.0,0.5,8,0.5,0,3.5,3100,5,0.01,810,0.10,40,10,0.01,0.20,0,0,0,0.01,0.30,0.08,0.50,0,None,0.10,0.05,"USDA-adaptado"),
    ("Caseína de potássio","Leite e derivados",6.0,368,1540,84.5,0.8,10,1.0,0,6.5,1700,15,0.01,900,0.15,10,300,0.02,0.50,0,0,0,0.01,0.30,0.08,0.50,0,None,0.12,0.05,"USDA-adaptado"),
    ("Lactoferrina bovina","Leite e derivados",5.0,450,1883,90.0,1.0,0,0,0,4.0,100,10,0.01,100,2.50,50,100,0.02,0.50,0,0,0,0,0.20,0,0.10,0,None,0,0,"USDA-adaptado"),
    ("Imunoglobulinas bovinas","Leite e derivados",5.0,400,1674,88.0,1.5,0,1.0,0,3.5,50,8,0.01,80,0.30,60,80,0.02,0.30,0,0,0,0,0.10,0,0.10,0,None,0,0,"USDA-adaptado"),
    ("Gordura de coco fracionada (MCT)","Óleos e gorduras",0.5,848,3549,0,99.5,0,0,0,0,1,0,None,0,0,0,0,None,None,0,0,0,0,0,0,0,0,None,0.50,0,"USDA-adaptado"),
    ("Proteína de ervilha isolada","Leguminosas",6.5,357,1494,83.0,6.5,0,2.5,1.0,2.0,45,120,0.55,700,9.50,750,280,0.35,2.80,0,0,0,0.50,0.20,0.30,1.50,0,None,0.50,0,"USDA-adaptado"),
    ("Proteína de arroz","Cereais e derivados",8.0,378,1582,80.0,7.0,0,3.0,1.0,2.0,40,80,3.50,850,4.20,50,280,0.30,3.50,0,0,0,0.30,0.10,0.35,4.50,0,None,0.90,0,"USDA-adaptado"),
    ("Extrato de levedura","Aditivos e ingredientes",4.0,325,1360,42.0,3.5,0,40.0,9.0,10.0,60,60,0.50,750,5.00,7000,1200,0.15,5.00,0,0,0,1.50,3.00,0.80,18.00,0,None,0,0,"USDA-adaptado"),
]


# ─── Suplemento 2: expansão adicional de alimentos ────────────────────────────
# Formato: (nome, categoria, umidade, kcal, kj, prot, lip, col, carb, fibra, cinzas,
#           ca, mg, mn, p, fe, na, k, cu, zn, retinol, re, rae, tia, ribo, pirid,
#           niac, vitc, vitd, vite, vitb12, fonte)
TACO_SUPLEMENTO_2 = [
    # ─── EMBUTIDOS E CARNES PROCESSADAS ──────────────────────────────────────
    ("Salame italiano","Carnes e derivados",28.5,431,1803,22.6,37.3,100,1.8,0,4.9,22,25,0.01,185,2.24,1740,320,0.15,3.62,0,0,0,0.64,0.25,0.39,5.80,0,None,0.31,0.90,"USDA-adaptado"),
    ("Linguiça calabresa","Carnes e derivados",52.0,270,1130,13.5,22.8,75,1.5,0,3.2,16,16,0.01,130,1.40,1150,230,0.09,2.40,0,0,0,0.35,0.18,0.22,3.80,0,None,0.28,0.60,"TBCA-USP"),
    ("Linguiça de frango","Carnes e derivados",60.0,170,711,16.0,11.0,68,1.0,0,2.5,14,20,0.01,140,0.90,850,220,0.07,1.80,0,0,0,0.15,0.12,0.25,4.50,0,None,0.22,0.35,"TBCA-USP"),
    ("Mortadela","Carnes e derivados",55.0,200,837,12.0,15.5,55,3.5,0,3.0,30,14,0.02,100,1.50,820,180,0.09,1.80,0,0,0,0.25,0.15,0.10,3.20,0,None,0.18,0.55,"TBCA-USP"),
    ("Presunto cozido","Carnes e derivados",67.0,130,544,17.5,5.5,57,1.5,0,2.2,10,18,0.01,185,1.10,900,250,0.09,2.10,0,0,0,0.52,0.18,0.35,3.90,0,None,0.25,0.52,"USDA-adaptado"),
    ("Peito de peru defumado","Carnes e derivados",72.0,109,456,17.8,2.0,50,2.5,0,2.0,12,22,0.01,175,0.95,860,290,0.07,1.80,0,0,0,0.04,0.12,0.32,5.80,0,None,0.20,0.40,"USDA-adaptado"),
    ("Bacon","Carnes e derivados",37.5,458,1917,12.6,41.5,110,1.5,0,2.5,6,14,0.01,135,1.16,1717,232,0.07,1.86,0,0,0,0.41,0.12,0.22,4.60,0,None,0.18,0.74,"USDA-adaptado"),
    ("Salsicha de frango","Carnes e derivados",58.0,210,879,12.0,17.0,65,2.0,0,2.5,30,16,0.01,110,1.20,700,190,0.08,1.60,0,0,0,0.12,0.10,0.20,3.50,0,None,0.20,0.38,"TBCA-USP"),
    ("Salsicha bovina","Carnes e derivados",55.0,243,1017,11.5,20.5,72,2.8,0,3.0,12,12,0.01,85,1.70,780,145,0.08,1.90,0,0,0,0.10,0.12,0.17,2.60,0,None,0.22,0.62,"TBCA-USP"),
    ("Copa lombo","Carnes e derivados",43.0,305,1276,25.5,21.5,80,0.5,0,3.5,18,22,0.01,210,1.80,1200,310,0.10,2.80,0,0,0,0.58,0.24,0.45,5.20,0,None,0.32,0.80,"USDA-adaptado"),
    ("Pastrami bovino","Carnes e derivados",60.5,147,615,21.0,5.5,68,2.5,0,2.5,12,25,0.01,195,2.80,1020,290,0.08,4.20,0,0,0,0.06,0.20,0.35,4.50,0,None,0.38,2.00,"USDA-adaptado"),
    ("Frango desfiado cozido","Carnes e derivados",70.0,141,590,25.0,4.0,80,0,0,1.0,12,27,0.02,210,0.70,60,280,0.05,1.60,0,0,0,0.06,0.13,0.68,11.90,0,None,0.27,0.35,"TBCA-USP"),
    ("Carne moída bovina 80%","Carnes e derivados",60.0,215,900,20.0,14.0,75,0,0,1.0,14,19,0.01,155,2.30,68,265,0.07,4.20,0,0,0,0.05,0.17,0.30,4.10,0,None,0.35,2.00,"TBCA-USP"),
    ("Peito de frango grelhado","Carnes e derivados",72.0,120,502,24.5,2.5,73,0,0,1.0,5,28,0.02,230,0.68,65,335,0.04,0.85,0,0,0,0.07,0.14,0.67,11.20,0,None,0.28,0.35,"TBCA-USP"),
    # ─── QUEIJOS ADICIONAIS ───────────────────────────────────────────────────
    ("Queijo feta","Leite e derivados",52.0,264,1105,14.2,21.3,89,4.1,0,5.0,493,19,0.01,337,0.65,1116,62,0.03,2.88,422,422,127,0.15,0.84,0.42,0.91,0,None,0.43,1.69,"USDA-adaptado"),
    ("Queijo halloumi","Leite e derivados",54.0,321,1343,21.4,25.0,76,1.9,0,2.5,700,20,0.01,390,0.30,1500,110,0.02,2.80,180,180,54,0.03,0.35,0.06,0.07,0,None,0.25,1.00,"USDA-adaptado"),
    ("Queijo mascarpone","Leite e derivados",44.0,429,1795,4.7,45.0,140,2.8,0,3.6,95,8,0.01,75,0.09,105,135,0.01,0.33,413,413,124,0.02,0.12,0.04,0.06,0,None,1.00,0.20,"USDA-adaptado"),
    ("Queijo burrata","Leite e derivados",58.0,250,1046,9.0,21.5,72,1.8,0,2.8,420,14,0.01,290,0.30,220,110,0.02,1.60,185,185,56,0.03,0.25,0.06,0.07,0,None,0.28,0.85,"USDA-adaptado"),
    ("Queijo de cabra fresco","Leite e derivados",63.5,198,829,10.9,15.1,40,2.5,0,2.5,140,16,0.03,256,1.62,368,158,0.36,0.63,205,205,62,0.05,0.21,0.07,0.27,0,None,0.19,0.07,"USDA-adaptado"),
    ("Queijo manchego","Leite e derivados",37.9,395,1653,26.4,32.0,93,1.2,0,3.0,748,32,0.01,540,0.28,900,142,0.04,4.60,310,310,93,0.04,0.46,0.08,0.10,0,None,0.28,1.80,"USDA-adaptado"),
    ("Queijo gruyère","Leite e derivados",33.2,413,1729,29.8,32.3,110,0.4,0,3.8,1011,36,0.02,605,0.17,336,81,0.03,3.90,295,295,89,0.06,0.28,0.08,0.14,0,None,0.23,3.34,"USDA-adaptado"),
    ("Queijo de ovelha","Leite e derivados",40.0,370,1549,22.0,30.0,80,1.5,0,4.0,680,26,0.01,450,0.35,800,120,0.03,3.50,270,270,81,0.03,0.42,0.07,0.08,0,None,0.24,1.40,"USDA-adaptado"),
    ("Queijo Philadelphia light","Leite e derivados",64.0,154,644,9.1,11.5,38,3.9,0,3.2,70,10,0.01,120,0.14,500,130,0.04,0.60,196,196,59,0.02,0.12,0.04,0.06,0,None,0.32,0.18,"USDA-adaptado"),
    ("Queijo mozzarella búfala","Leite e derivados",57.5,256,1071,17.0,20.0,52,2.5,0,2.2,400,20,0.01,310,0.28,270,115,0.02,1.90,200,200,60,0.03,0.28,0.06,0.07,0,None,0.25,0.96,"USDA-adaptado"),
    # ─── LATICÍNIOS FUNCIONAIS E ESPECIALIZADOS ───────────────────────────────
    ("Leite de búfala integral","Leite e derivados",82.0,97,406,4.5,6.9,19,5.2,0,0.8,195,31,0.01,117,0.12,52,178,0.05,0.22,52,52,16,0.05,0.13,0.04,0.15,2.3,None,0.10,0.36,"USDA-adaptado"),
    ("Leite de cabra integral","Leite e derivados",87.0,69,289,3.6,4.1,11,4.5,0,0.8,134,14,0.02,111,0.05,50,204,0.04,0.30,56,56,17,0.05,0.14,0.07,0.28,1.3,None,0.07,0.07,"USDA-adaptado"),
    ("Leite de ovelha integral","Leite e derivados",81.0,108,452,5.5,7.0,27,5.4,0,0.9,194,18,0.01,158,0.10,44,137,0.04,0.57,72,72,22,0.07,0.35,0.08,0.42,1.8,None,0.11,0.71,"USDA-adaptado"),
    ("Coalhada seca","Leite e derivados",72.0,90,377,7.5,3.8,12,8.5,0,1.2,145,14,0.01,112,0.05,65,190,0.01,0.58,24,24,7,0.04,0.20,0.05,0.11,0.9,None,0.06,0.42,"TBCA-USP"),
    ("Coalhada fresca","Leite e derivados",85.0,52,218,4.2,1.5,7,7.5,0,1.0,130,13,0.01,102,0.04,58,175,0.01,0.52,20,20,6,0.04,0.18,0.05,0.09,0.8,None,0.05,0.38,"TBCA-USP"),
    ("Skyr (proteína islandesa)","Leite e derivados",82.0,63,264,11.0,0.2,3,4.5,0,0.8,135,14,0.01,120,0.05,55,200,0.01,0.60,8,8,2,0.04,0.25,0.05,0.15,0.5,None,0.02,1.00,"USDA-adaptado"),
    ("Manteiga de búfala","Óleos e gorduras",14.0,745,3117,0.5,83.0,240,0.1,0,2.5,16,2,0.00,20,0.05,25,30,0.00,0.12,800,800,800,0.00,0.04,0.00,0.05,0,None,2.60,0.20,"USDA-adaptado"),
    ("Queijo fundido para molho","Leite e derivados",47.0,302,1264,12.5,24.5,72,9.5,0,5.5,620,20,0.01,450,0.25,1120,130,0.02,2.40,195,195,58,0.03,0.30,0.06,0.07,0,None,0.28,0.78,"TBCA-USP"),
    ("Ricota defumada","Leite e derivados",60.5,204,854,12.8,16.0,60,2.8,0,1.5,230,12,0.01,175,0.26,310,118,0.03,1.30,130,130,39,0.01,0.22,0.05,0.10,0,None,0.45,0.40,"TBCA-USP"),
    ("Nata (sour cream)","Leite e derivados",71.0,193,808,2.8,19.4,52,3.5,0,0.6,106,10,0.01,72,0.04,53,148,0.01,0.28,188,188,56,0.03,0.14,0.03,0.07,0.4,None,0.72,0.23,"USDA-adaptado"),
    ("Chantilly UHT","Leite e derivados",60.0,300,1255,2.0,30.0,88,4.5,0,0.5,80,8,0.01,60,0.03,30,110,0.01,0.32,305,305,91,0.02,0.12,0.03,0.06,0.4,None,1.40,0.18,"TBCA-USP"),
    ("Leite de coco integral","Bebidas",68.0,197,824,2.0,21.3,0,2.8,2.2,0.7,16,37,0.92,100,1.64,13,263,0.26,0.67,0,0,0,0.03,0.00,0.03,0.76,2.8,None,0.15,0,"USDA-adaptado"),
    ("Creme de coco","Óleos e gorduras",32.0,330,1381,3.6,34.7,0,6.7,2.8,0.8,11,28,0.55,60,3.32,18,250,0.25,0.72,0,0,0,0.02,0.01,0.03,0.80,1.0,None,0.17,0,"USDA-adaptado"),
    # ─── VEGETAIS ADICIONAIS ──────────────────────────────────────────────────
    ("Aspargo cru","Hortaliças",93.2,20,84,2.2,0.1,0,3.9,2.1,0.6,24,14,0.16,52,2.14,2,202,0.19,0.54,38,38,25,0.14,0.13,0.09,0.98,5.6,None,1.13,0,"USDA-adaptado"),
    ("Alcachofra","Hortaliças",84.9,47,197,3.3,0.2,0,10.5,5.4,1.0,44,60,0.26,90,1.28,94,370,0.23,0.49,1,1,0,0.07,0.07,0.12,1.05,11.7,None,0.19,0,"USDA-adaptado"),
    ("Cogumelo shitake","Hortaliças",89.7,34,142,2.2,0.5,0,6.8,2.5,0.8,2,20,0.23,112,0.41,9,304,0.14,1.03,0,0,0,0.02,0.22,0.30,3.88,0,None,0.01,0,"USDA-adaptado"),
    ("Cogumelo champignon","Hortaliças",92.1,22,92,3.1,0.3,0,3.3,1.0,1.0,3,9,0.05,86,0.50,5,318,0.30,0.52,0,0,0,0.08,0.40,0.10,3.60,2.1,None,0.01,0,"TACO 4ª edição"),
    ("Palmito em conserva","Hortaliças",90.7,28,117,2.7,0.4,0,3.4,2.2,0.8,12,30,0.32,28,1.46,396,300,0.18,0.22,0,0,0,0.02,0.03,0.05,0.50,0.5,None,0.56,0,"TACO 4ª edição"),
    ("Pimentão amarelo","Hortaliças",92.2,27,113,1.0,0.2,0,6.3,0.9,0.5,11,12,0.11,24,0.46,2,212,0.06,0.17,20,20,10,0.05,0.03,0.17,0.89,184.0,None,0.59,0,"USDA-adaptado"),
    ("Alho poró","Hortaliças",83.0,61,255,1.5,0.3,0,14.1,1.8,0.9,59,28,0.48,35,2.10,20,180,0.12,0.12,83,83,64,0.06,0.03,0.23,0.40,12.0,None,0.92,0,"USDA-adaptado"),
    ("Nabo","Hortaliças",91.9,28,117,0.9,0.1,0,6.4,1.8,0.6,30,11,0.13,27,0.30,67,191,0.09,0.27,0,0,0,0.04,0.03,0.09,0.40,21.0,None,0.03,0,"USDA-adaptado"),
    ("Raiz-forte (wasabi)","Condimentos",77.5,109,456,4.5,0.7,0,23.4,7.8,0.9,100,69,0.15,80,1.03,17,246,0.06,0.82,2,2,1,0.13,0.11,0.22,0.39,24.9,None,0.13,0,"USDA-adaptado"),
    ("Coentro fresco","Temperos e molhos",92.2,23,96,2.1,0.5,0,3.7,2.8,1.5,67,26,0.43,48,1.77,46,521,0.22,0.50,337,337,269,0.07,0.16,0.15,1.11,27.0,None,2.50,0,"USDA-adaptado"),
    ("Hortelã fresca","Temperos e molhos",85.6,70,293,3.7,0.9,0,14.9,8.0,1.8,243,80,1.18,73,5.08,31,569,0.24,1.11,212,212,203,0.08,0.27,0.16,1.71,31.8,None,5.08,0,"USDA-adaptado"),
    ("Manjericão fresco","Temperos e molhos",92.1,23,96,3.2,0.6,0,2.7,1.6,1.5,177,64,1.15,56,3.17,4,295,0.38,0.81,264,264,264,0.03,0.08,0.16,0.90,18.0,None,0.80,0,"USDA-adaptado"),
    ("Alecrim fresco","Temperos e molhos",67.8,131,548,3.3,5.9,0,20.7,14.1,2.8,317,91,0.96,66,6.65,26,668,0.30,0.93,146,146,146,0.04,0.19,0.34,0.91,21.8,None,3.53,0,"USDA-adaptado"),
    ("Tomilho fresco","Temperos e molhos",65.1,101,423,5.6,1.7,0,24.5,14.0,4.0,405,160,1.72,106,17.45,9,609,0.55,1.81,238,238,238,0.05,0.47,0.35,1.82,160.1,None,7.43,0,"USDA-adaptado"),
    ("Louro folha seca","Especiarias e condimentos",5.4,313,1310,7.6,8.4,0,74.0,26.3,4.2,834,120,8.17,113,43.00,23,529,0.42,3.70,309,309,0,0.01,0.42,1.74,2.01,46.5,None,5.10,0,"USDA-adaptado"),
    ("Pimenta do reino moída","Especiarias e condimentos",11.4,251,1050,10.4,3.3,0,63.9,25.3,3.1,443,171,12.75,173,9.71,20,1329,1.33,1.42,27,27,15,0.11,0.18,0.34,1.14,0,None,1.04,0,"USDA-adaptado"),
    ("Cominho em pó","Especiarias e condimentos",8.1,375,1569,17.8,22.3,0,44.2,10.5,5.6,931,366,3.33,499,66.36,168,1788,0.87,5.50,64,64,27,0.63,0.33,0.44,4.58,7.7,None,3.33,0,"USDA-adaptado"),
    ("Orégano seco","Especiarias e condimentos",9.9,265,1109,9.0,4.3,0,68.9,42.5,7.6,1597,270,4.99,148,36.80,25,1260,0.63,2.69,690,690,85,0.18,0.29,1.04,4.64,2.3,None,18.26,0,"USDA-adaptado"),
    ("Canela em pó","Especiarias e condimentos",10.6,247,1034,4.0,1.2,0,80.6,53.1,1.5,1002,60,17.47,64,8.32,10,431,0.34,1.83,15,15,0,0.02,0.04,0.04,1.33,3.8,None,2.32,0,"USDA-adaptado"),
    ("Páprica doce","Especiarias e condimentos",11.2,282,1180,14.1,12.9,0,53.9,34.9,7.0,229,178,1.59,314,21.14,68,2280,0.71,4.33,2463,2463,949,0.33,1.23,3.14,10.06,0,None,29.10,0,"USDA-adaptado"),
    ("Noz moscada moída","Especiarias e condimentos",6.2,525,2197,5.8,36.3,0,49.3,20.8,2.5,184,183,2.90,213,3.04,16,350,1.03,2.15,5,5,0,0.35,0.06,0.16,1.35,3.0,None,0.19,0,"USDA-adaptado"),
    ("Cravo em pó","Especiarias e condimentos",6.9,274,1147,5.9,13.0,0,65.5,33.9,6.5,646,259,60.11,104,11.83,277,1102,0.37,2.32,27,27,0,0.16,0.22,0.39,1.56,80.8,None,8.82,0,"USDA-adaptado"),
    ("Gengibre em pó","Especiarias e condimentos",9.9,335,1402,8.98,4.24,0,71.62,14.1,4.8,114,214,33.30,168,19.80,27,1320,0.48,3.64,30,30,0,0.046,0.17,1.03,9.62,0,None,0,0,"USDA-adaptado"),
    ("Curry em pó","Especiarias e condimentos",9.6,325,1360,12.7,14.0,0,55.8,33.2,7.2,478,254,9.45,349,29.59,52,1543,0.98,4.05,436,436,96,0.24,0.27,0.79,4.21,10.1,None,25.24,0,"USDA-adaptado"),
    # ─── MOLHOS, PATÊS E CONDIMENTOS ─────────────────────────────────────────
    ("Maionese tradicional","Óleos e gorduras",15.0,680,2845,1.1,74.9,63,3.1,0,1.0,14,2,0.01,18,0.22,760,28,0.01,0.28,72,72,22,0.01,0.05,0.01,0.07,0,None,9.32,0.07,"USDA-adaptado"),
    ("Maionese light","Óleos e gorduras",34.0,335,1402,1.5,32.5,41,7.5,0,1.0,22,3,0.01,22,0.15,620,40,0.02,0.30,58,58,17,0.01,0.04,0.01,0.05,0,None,5.50,0.05,"USDA-adaptado"),
    ("Molho shoyu","Temperos e molhos",71.0,60,251,5.8,0.1,0,5.6,0.6,14.0,18,45,0.35,103,2.41,5717,341,0.12,0.42,0,0,0,0.07,0.22,0.32,2.10,0,None,0,0,"USDA-adaptado"),
    ("Molho inglês (worcestershire)","Temperos e molhos",75.0,78,326,2.1,0.1,0,19.5,0.4,3.1,78,27,1.37,45,3.96,980,345,0.24,0.26,0,0,0,0.17,0.07,0.01,1.14,17.1,None,0.83,0,"USDA-adaptado"),
    ("Extrato de tomate","Temperos e molhos",68.6,83,347,4.5,0.4,0,19.0,4.2,3.6,46,52,0.57,103,2.60,987,1014,0.39,0.53,338,338,202,0.14,0.15,0.22,3.61,28.3,None,4.39,0,"USDA-adaptado"),
    ("Mostarda dijon","Temperos e molhos",77.0,98,410,6.1,5.1,0,9.7,3.7,2.5,93,38,0.32,94,1.74,1104,162,0.30,0.67,8,8,4,0.12,0.09,0.17,0.97,0.3,None,0.36,0,"USDA-adaptado"),
    ("Molho de pimenta tabasco","Temperos e molhos",88.0,30,126,0.9,0.5,0,6.0,0.8,3.0,6,4,0.08,14,0.50,3067,155,0.09,0.10,1862,1862,1117,0.02,0.06,0.10,0.30,0,None,0.16,0,"USDA-adaptado"),
    ("Tahine escuro","Oleaginosas",4.0,592,2477,17.0,52.7,0,21.2,6.4,3.1,620,95,1.05,573,8.95,115,414,1.49,4.62,0,0,0,0.24,0.47,0.15,5.45,0,None,0.25,0,"USDA-adaptado"),
    ("Pasta de amendoim integral","Oleaginosas",1.2,588,2461,24.1,50.4,0,18.6,5.9,3.4,55,185,1.97,363,1.80,18,639,0.44,2.80,0,0,0,0.12,0.14,0.44,14.08,0,None,7.24,0,"USDA-adaptado"),
    # ─── BEBIDAS E SUCOS ──────────────────────────────────────────────────────
    ("Kombucha natural","Bebidas",97.0,13,54,0.1,0,0,3.0,0,0.1,5,2,0.02,5,0.02,5,30,0.01,0.02,0,0,0,0.02,0.04,0.02,0.06,0,None,0,0,"USDA-adaptado"),
    ("Kefir de água","Bebidas",97.5,17,71,0.2,0,0,4.0,0,0.1,6,3,0.02,8,0.03,10,35,0.01,0.03,0,0,0,0.03,0.06,0.03,0.10,0,None,0,0,"USDA-adaptado"),
    ("Suco de maracujá natural","Bebidas",88.0,44,184,0.8,0.1,0,10.0,0.3,0.4,4,10,0.05,10,0.11,4,140,0.06,0.05,60,60,36,0.00,0.10,0.07,1.50,20.0,None,0.05,0,"TBCA-USP"),
    ("Suco de acerola","Bebidas",90.5,40,167,0.7,0.2,0,9.5,0.8,0.4,12,18,0.08,17,0.20,7,130,0.15,0.08,35,35,21,0.02,0.05,0.02,0.38,830.0,None,0.10,0,"TBCA-USP"),
    ("Suco de manga","Bebidas",84.7,62,259,0.8,0.2,0,14.8,0.6,0.4,12,14,0.05,14,0.20,4,170,0.07,0.06,358,358,215,0.03,0.03,0.06,0.48,22.0,None,0.80,0,"TBCA-USP"),
    ("Suco de abacaxi","Bebidas",87.0,46,193,0.5,0.1,0,11.2,0.4,0.3,14,14,0.45,7,0.22,1,130,0.05,0.07,2,2,1,0.07,0.02,0.08,0.45,20.0,None,0.02,0,"TBCA-USP"),
    ("Bebida isotônica","Bebidas",93.7,25,105,0,0,0,6.5,0,0.8,2,3,0.02,2,0.02,450,20,0.01,0.04,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Chá verde infusão","Bebidas",99.5,1,4,0.2,0,0,0.2,0,0,0,1,0.11,1,0.03,2,20,0.01,0.01,0,0,0,0.01,0.06,0,0,0,None,0,0,"USDA-adaptado"),
    ("Café coado filtrado","Bebidas",99.6,2,8,0.3,0,0,0,0,0.1,2,3,0.01,7,0.02,2,49,0.01,0.02,0,0,0,0,0.18,0,0.10,0,None,0,0,"USDA-adaptado"),
    ("Whey pronto para beber (RTD)","Bebidas",88.0,60,251,7.5,1.0,4,5.5,0,0.8,200,22,0.02,180,0.20,80,200,0.03,1.20,0,0,0,0.08,0.20,0.08,0.40,0,0.5,0.15,0.80,"USDA-adaptado"),
    ("Leite achocolatado UHT","Bebidas",82.0,78,326,3.5,2.5,12,12.0,0.3,0.8,120,20,0.40,110,0.40,55,200,0.09,0.80,45,45,13,0.08,0.26,0.08,0.28,0.3,None,0.60,0.50,"TBCA-USP"),
    # ─── FRUTAS ADICIONAIS ────────────────────────────────────────────────────
    ("Banana d'água","Frutas",73.5,101,423,1.2,0.2,0,25.0,2.2,0.7,5,28,0.30,28,0.36,2,358,0.12,0.17,8,8,5,0.04,0.06,0.35,0.60,9.0,None,0.14,0,"TACO 4ª edição"),
    ("Banana nanica","Frutas",74.8,92,385,1.4,0.1,0,23.0,1.9,0.8,3,26,0.32,22,0.28,1,362,0.09,0.14,4,4,2,0.03,0.05,0.37,0.64,9.1,None,0.12,0,"TACO 4ª edição"),
    ("Laranja bahia","Frutas",86.6,49,205,1.0,0.2,0,12.2,2.4,0.5,51,12,0.02,25,0.12,1,178,0.04,0.07,14,14,8,0.09,0.04,0.06,0.26,65.0,None,0.18,0,"TACO 4ª edição"),
    ("Laranja pera","Frutas",87.2,47,197,0.9,0.2,0,11.8,2.3,0.5,46,11,0.02,23,0.12,1,165,0.04,0.07,13,13,8,0.08,0.04,0.05,0.24,62.0,None,0.17,0,"TACO 4ª edição"),
    ("Limão galego","Frutas",91.4,30,126,0.7,0.3,0,6.6,1.8,0.4,30,8,0.03,15,0.24,1,138,0.04,0.11,2,2,1,0.04,0.02,0.06,0.12,50.0,None,0.20,0,"TBCA-USP"),
    ("Tangerina murcote","Frutas",88.1,39,163,0.9,0.2,0,9.7,1.8,0.4,30,11,0.05,14,0.14,1,157,0.04,0.05,16,16,10,0.09,0.04,0.06,0.28,36.0,None,0.20,0,"TACO 4ª edição"),
    ("Maracujá polpa","Frutas",79.5,68,285,2.0,1.3,0,14.3,1.3,1.2,13,28,0.08,68,1.60,28,348,0.09,0.10,64,64,39,0.00,0.20,0.10,1.50,30.0,None,0.02,0,"TACO 4ª edição"),
    ("Carambola","Frutas",91.4,31,130,1.0,0.3,0,6.9,2.8,0.5,3,10,0.04,12,0.08,2,133,0.07,0.12,3,3,2,0.01,0.02,0.02,0.36,34.4,None,0.15,0,"USDA-adaptado"),
    ("Jabuticaba","Frutas",81.6,58,243,0.6,0.1,0,14.7,0.4,0.3,9,18,0.08,10,0.32,2,130,0.04,0.08,2,2,1,0.04,0.08,0.04,0.40,21.6,None,0.19,0,"TBCA-USP"),
    ("Caju pedúnculo","Frutas",87.1,38,159,1.0,0.2,0,9.0,1.4,0.4,4,18,0.13,33,0.32,3,188,0.26,0.26,1,1,1,0.01,0.03,0.17,0.52,219.7,None,0.28,0,"TACO 4ª edição"),
    ("Umbu","Frutas",90.5,35,146,0.5,0.2,0,8.2,0.6,0.4,14,14,0.05,14,0.25,3,130,0.09,0.10,3,3,2,0.02,0.04,0.06,0.48,33.0,None,0.20,0,"TBCA-USP"),
    ("Framboesa","Frutas",85.8,52,218,1.2,0.7,0,11.9,6.5,0.6,25,22,0.67,29,0.69,1,151,0.09,0.42,2,2,1,0.03,0.04,0.06,0.60,26.2,None,0.87,0,"USDA-adaptado"),
    ("Mirtilo (blueberry)","Frutas",84.2,57,239,0.7,0.3,0,14.5,2.4,0.2,6,6,0.34,12,0.28,1,77,0.06,0.16,3,3,2,0.04,0.04,0.05,0.42,9.7,None,0.57,0,"USDA-adaptado"),
    ("Amora preta","Frutas",88.2,43,180,1.4,0.5,0,9.6,5.3,0.4,29,20,0.65,22,0.62,1,162,0.17,0.53,11,11,7,0.02,0.04,0.03,0.65,21.0,None,1.17,0,"USDA-adaptado"),
    ("Romã","Frutas",77.9,83,347,1.7,1.2,0,18.7,4.0,0.7,10,12,0.12,36,0.30,3,236,0.16,0.35,0,0,0,0.07,0.05,0.08,0.29,10.2,None,0.60,0,"USDA-adaptado"),
    ("Caqui","Frutas",80.3,70,293,0.6,0.2,0,18.6,3.6,0.5,8,9,0.36,17,0.15,1,161,0.11,0.11,81,81,50,0.03,0.02,0.10,0.10,7.5,None,0.73,0,"USDA-adaptado"),
    ("Nectarina","Frutas",87.6,44,184,1.1,0.3,0,10.5,1.7,0.5,6,9,0.05,26,0.28,0,201,0.09,0.17,17,17,10,0.03,0.03,0.02,1.13,5.4,None,0.77,0,"USDA-adaptado"),
    ("Cereja","Frutas",81.8,63,264,1.1,0.2,0,16.0,2.1,0.5,13,11,0.07,21,0.36,0,222,0.06,0.07,3,3,2,0.03,0.04,0.05,0.20,7.0,None,0.07,0,"USDA-adaptado"),
    ("Damasco fresco","Frutas",86.4,48,201,1.4,0.4,0,11.1,2.0,0.8,13,10,0.08,23,0.39,1,259,0.08,0.20,96,96,58,0.03,0.04,0.05,0.60,10.0,None,0.89,0,"USDA-adaptado"),
    # ─── PANIFICAÇÃO E CONFEITARIA ────────────────────────────────────────────
    ("Pão francês","Cereais e derivados",36.3,300,1255,8.4,1.5,0,58.4,1.6,1.5,59,21,0.56,90,2.10,540,122,0.12,0.82,0,0,0,0.22,0.14,0.05,2.55,0,None,0.35,0,"TACO 4ª edição"),
    ("Pão de queijo","Cereais e derivados",35.0,279,1167,5.1,12.0,30,37.0,0.8,1.5,200,18,0.05,180,0.55,455,120,0.02,1.30,90,90,27,0.05,0.22,0.06,0.15,0.2,None,0.25,0.40,"TBCA-USP"),
    ("Croissant manteiga","Cereais e derivados",21.5,406,1699,8.1,21.3,52,46.0,1.5,2.0,40,15,0.38,90,2.40,400,120,0.08,0.65,138,138,41,0.25,0.17,0.06,2.40,0,None,0.65,0.28,"USDA-adaptado"),
    ("Biscoito cream cracker","Cereais e derivados",3.0,416,1741,8.0,11.0,0,73.0,2.5,2.5,32,17,0.56,110,2.50,590,135,0.12,0.70,0,0,0,0.15,0.10,0.06,2.10,0,None,0.40,0,"TBCA-USP"),
    ("Biscoito wafer","Cereais e derivados",2.8,488,2042,7.0,26.5,5,61.0,1.8,1.8,28,16,0.38,85,2.10,230,115,0.08,0.55,80,80,24,0.12,0.10,0.05,1.80,0,None,0.55,0.08,"TBCA-USP"),
    ("Panetone","Cereais e derivados",27.5,337,1410,8.2,10.5,95,54.0,2.0,1.5,40,15,0.35,90,2.00,350,130,0.09,0.70,80,80,24,0.20,0.14,0.06,2.20,0,None,0.65,0.25,"TBCA-USP"),
    ("Massa folhada","Cereais e derivados",17.0,464,1942,6.4,29.6,0,50.5,1.8,1.5,24,14,0.42,70,2.20,450,115,0.08,0.55,0,0,0,0.26,0.15,0.06,2.35,0,None,1.58,0,"USDA-adaptado"),
    ("Massa pizza pré-assada","Cereais e derivados",35.0,264,1105,8.0,3.2,0,50.5,2.1,1.8,48,18,0.50,85,1.95,410,115,0.10,0.72,0,0,0,0.22,0.13,0.05,2.40,0,None,0.35,0,"USDA-adaptado"),
    ("Farinha láctea","Cereais e derivados",5.5,405,1694,16.8,9.2,38,68.0,3.5,3.5,500,55,0.25,420,3.50,300,750,0.07,2.50,320,320,96,0.40,1.00,0.30,3.50,12.0,None,0.60,2.00,"TBCA-USP"),
    ("Creme de confeiteiro","Leite e derivados",60.0,180,753,4.5,7.5,68,25.0,0,0.8,120,13,0.01,115,0.28,80,175,0.01,0.60,78,78,23,0.07,0.25,0.06,0.18,0.5,None,0.32,0.38,"TBCA-USP"),
    ("Ganache escuro 60%","Outros",15.5,485,2030,4.5,33.5,5,48.5,5.0,2.0,40,112,0.95,148,5.80,10,380,0.85,1.60,0,0,0,0.03,0.05,0.04,0.52,0,None,0.90,0,"TBCA-USP"),
    ("Cobertura de chocolate hidrogenada","Outros",2.0,560,2344,3.5,38.0,0,55.5,3.5,0.8,30,50,0.25,80,2.50,12,280,0.40,0.80,0,0,0,0.02,0.04,0.02,0.40,0,None,0.60,0,"USDA-adaptado"),
    ("Fondant de açúcar","Açúcares e produtos",8.0,357,1494,0,0,0,90.5,0,0.5,3,1,0.01,2,0.05,5,8,0.01,0.01,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Açúcar impalpável","Açúcares e produtos",0.5,395,1653,0,0,0,99.0,0,0.1,1,0,0,0,0.02,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Glucose de arroz","Cereais e derivados",18.0,308,1289,0.1,0.1,0,81.5,0.1,0.3,2,1,0.01,4,0.05,8,8,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Xarope de agave","Açúcares e produtos",24.0,310,1297,0.1,0.5,0,76.0,0,0,3,1,0.04,3,0.05,5,4,0.04,0.02,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Xarope de bordo (maple)","Açúcares e produtos",32.4,260,1088,0,0.1,0,67.1,0,0.5,102,14,2.29,2,0.11,9,212,0.06,1.13,0,0,0,0.01,0.65,0,0.08,0,None,0,0,"USDA-adaptado"),
    # ─── CEREAIS E GRÃOS ADICIONAIS ───────────────────────────────────────────
    ("Trigo sarraceno","Cereais e derivados",9.8,343,1435,13.2,3.4,0,71.5,10.0,2.1,18,231,1.30,347,2.20,1,460,1.10,2.40,0,0,0,0.10,0.43,0.21,7.02,0,None,0.80,0,"USDA-adaptado"),
    ("Amaranto grão","Cereais e derivados",11.3,371,1553,13.6,7.0,0,65.3,6.7,2.9,159,248,2.26,557,7.61,4,508,0.53,2.87,0,0,0,0.12,0.20,0.59,1.29,4.2,None,1.19,0,"USDA-adaptado"),
    ("Teff grão","Cereais e derivados",11.0,367,1536,13.3,2.4,0,73.1,8.0,2.0,180,184,9.24,429,7.63,12,427,0.81,3.63,0,0,0,0.39,0.27,0.48,3.36,0,None,0.90,0,"USDA-adaptado"),
    ("Sorgo grão","Cereais e derivados",9.2,339,1419,11.3,3.3,0,74.6,6.3,1.6,28,165,1.56,287,4.40,6,350,0.28,1.67,0,0,0,0.24,0.14,0.44,2.93,0,None,0.50,0,"USDA-adaptado"),
    ("Milho de pipoca cru","Cereais e derivados",12.0,375,1569,11.0,5.0,0,74.0,14.5,1.2,12,137,0.83,358,2.71,8,329,0.27,2.72,11,11,6,0.27,0.12,0.33,1.83,0,None,0.49,0,"USDA-adaptado"),
    ("Trigo burgul cozido","Cereais e derivados",78.0,83,347,3.1,0.2,0,18.6,4.5,0.7,10,32,0.70,40,0.96,5,68,0.10,0.57,0,0,0,0.10,0.03,0.08,1.30,0,None,0.10,0,"USDA-adaptado"),
    ("Cuscuz marroquino cozido","Cereais e derivados",72.6,112,469,3.8,0.2,0,23.2,1.4,0.4,7,8,0.11,24,0.26,6,58,0.05,0.26,0,0,0,0.06,0.04,0.05,1.20,0,None,0.09,0,"USDA-adaptado"),
    ("Farelo de trigo","Cereais e derivados",9.9,216,904,15.5,4.3,0,64.5,42.8,4.1,73,611,11.50,1013,10.57,2,1182,0.99,7.27,0,0,0,0.52,0.58,1.30,13.58,0,None,1.49,0,"USDA-adaptado"),
    ("Farelo de aveia","Cereais e derivados",6.6,246,1030,17.3,7.0,0,66.2,15.4,3.5,58,235,5.63,734,5.41,4,566,0.67,5.12,0,0,0,0.82,0.28,0.19,1.20,0,None,1.01,0,"USDA-adaptado"),
    ("Arroz integral cozido","Cereais e derivados",73.9,124,519,2.6,1.0,0,25.6,1.8,0.7,10,44,1.10,83,0.42,5,79,0.16,0.62,0,0,0,0.18,0.02,0.15,1.90,0,None,0.15,0,"USDA-adaptado"),
    ("Macarrão integral cozido","Cereais e derivados",68.0,130,544,5.3,1.1,0,27.2,3.9,0.6,15,31,1.20,91,1.36,4,75,0.29,1.02,0,0,0,0.12,0.05,0.08,1.65,0,None,0.40,0,"USDA-adaptado"),
    ("Granola sem adição de açúcar","Cereais e derivados",5.0,420,1758,11.5,15.0,0,60.0,8.5,3.0,52,90,2.50,300,3.30,25,285,0.38,2.80,0,0,0,0.38,0.12,0.12,2.60,0.4,None,5.00,0,"USDA-adaptado"),
    ("Flocos de milho (cornflakes)","Cereais e derivados",3.0,357,1494,7.5,0.7,0,84.0,3.0,2.5,5,8,0.24,82,8.33,659,96,0.08,0.45,150,150,45,0.90,1.00,1.14,11.90,15.0,4.54,0.16,1.43,"USDA-adaptado"),
    # ─── LEGUMINOSAS E PROTEÍNAS VEGETAIS ────────────────────────────────────
    ("Feijão branco cozido","Leguminosas",67.0,139,582,9.7,0.5,0,25.1,6.3,1.4,90,63,0.51,177,3.70,2,561,0.24,1.28,0,0,0,0.24,0.08,0.13,0.43,0,None,0.20,0,"USDA-adaptado"),
    ("Feijão-de-corda cozido","Leguminosas",71.0,101,423,7.4,0.4,0,18.3,4.5,0.8,25,53,0.28,147,2.13,5,251,0.19,1.09,0,0,0,0.16,0.04,0.06,0.65,0.5,None,0.06,0,"TACO 4ª edição"),
    ("Ervilha enlatada","Leguminosas",76.0,72,301,4.5,0.3,0,12.8,4.4,0.8,24,22,0.26,65,1.29,291,148,0.16,0.78,28,28,18,0.11,0.05,0.08,0.90,8.0,None,0.06,0,"USDA-adaptado"),
    ("Grão-de-bico cru","Leguminosas",11.5,364,1523,19.3,6.0,0,61.0,17.4,3.0,105,115,2.20,366,6.24,24,875,0.85,3.43,3,3,2,0.48,0.21,0.54,1.54,4.0,None,0.82,0,"USDA-adaptado"),
    ("Lentilha vermelha crua","Leguminosas",9.5,352,1473,25.8,1.1,0,60.1,10.8,2.0,51,122,1.33,451,7.54,6,955,0.76,3.27,2,2,2,0.38,0.21,0.54,2.60,1.5,None,0.49,0,"USDA-adaptado"),
    ("Edamame cozido","Leguminosas",67.5,122,511,11.9,5.2,0,8.9,5.2,1.3,63,64,1.02,161,2.27,6,436,0.37,1.64,9,9,5,0.30,0.18,0.10,1.06,6.1,None,0.37,0,"USDA-adaptado"),
    ("Tofu firme","Leguminosas",76.2,76,318,8.1,4.8,0,1.9,0.3,0.5,350,30,0.61,97,5.36,7,121,0.19,1.02,1,1,1,0.08,0.04,0.05,0.20,0.1,None,0.01,0,"USDA-adaptado"),
    ("Tempê","Leguminosas",59.6,195,816,20.3,10.8,0,7.6,0,2.2,111,81,1.30,266,2.70,9,412,0.57,1.73,0,0,0,0.08,0.36,0.24,2.64,0,None,0.61,0,"USDA-adaptado"),
    ("Missô (pasta de soja)","Leguminosas",43.0,199,833,11.7,6.0,0,26.5,5.4,12.7,57,48,0.82,159,2.49,3728,210,0.41,2.56,9,9,5,0.10,0.23,0.19,1.08,0,None,0.21,0.08,"USDA-adaptado"),
    # ─── PEIXES E FRUTOS DO MAR ADICIONAIS ───────────────────────────────────
    ("Bacalhau seco salgado","Peixes e frutos do mar",19.6,280,1172,62.8,2.3,100,0,0,15.2,62,67,0.05,525,1.31,7027,908,0.22,1.59,0,0,0,0.14,0.09,0.36,6.51,0,None,0.78,2.61,"USDA-adaptado"),
    ("Bacalhau dessalgado cozido","Peixes e frutos do mar",72.5,105,439,25.1,0.9,62,0,0,1.4,28,35,0.02,270,0.60,2700,450,0.10,0.80,0,0,0,0.07,0.05,0.18,3.50,0,None,0.40,1.50,"TBCA-USP"),
    ("Camarão fresco cru","Peixes e frutos do mar",78.5,85,356,18.0,1.0,126,0,0,1.5,52,22,0.05,223,2.65,119,259,0.22,1.34,54,54,13,0.02,0.01,0.20,2.55,1.6,None,1.32,1.11,"USDA-adaptado"),
    ("Polvo cozido","Peixes e frutos do mar",72.0,164,686,29.8,2.1,96,4.4,0,1.8,106,51,0.11,454,8.13,311,630,0.37,3.34,10,10,2,0.04,0.07,0.65,3.73,7.8,None,1.20,36.00,"USDA-adaptado"),
    ("Lula cozida","Peixes e frutos do mar",73.0,175,733,32.0,2.7,233,3.3,0,1.3,44,33,0.07,365,0.68,260,279,1.89,1.53,0,0,0,0.06,0.55,0.23,2.97,4.2,None,1.20,1.30,"USDA-adaptado"),
    ("Ostra cozida","Peixes e frutos do mar",74.0,163,682,18.9,5.7,89,9.9,0,2.3,97,47,5.84,333,6.66,220,369,4.49,90.81,0,0,0,0.14,0.23,0.10,3.00,5.0,None,0.98,16.62,"USDA-adaptado"),
    ("Mexilhão cozido","Peixes e frutos do mar",58.0,172,720,23.8,4.5,56,7.4,0,2.3,33,37,6.80,285,6.72,369,320,0.14,2.67,48,48,13,0.25,0.36,0.08,2.55,12.0,None,0.98,12.00,"USDA-adaptado"),
    ("Salmão defumado","Peixes e frutos do mar",72.0,117,490,18.3,4.3,23,0,0,2.9,11,18,0.02,186,0.85,784,175,0.17,0.31,27,27,7,0.18,0.16,0.49,5.00,0,5.10,1.02,3.26,"USDA-adaptado"),
    ("Cação (tubarão) cozido","Peixes e frutos do mar",74.0,130,544,21.3,4.5,51,0,0,1.3,22,22,0.02,218,0.98,79,350,0.07,0.65,0,0,0,0.07,0.05,0.24,3.90,0,None,0.84,1.10,"TBCA-USP"),
    # ─── OVOS E DERIVADOS ─────────────────────────────────────────────────────
    ("Ovo de pata","Ovos",70.8,185,774,13.0,13.7,884,1.4,0,1.1,64,16,0.04,220,3.85,146,156,0.09,1.41,383,383,383,0.26,0.46,0.26,0.32,0,3.77,1.32,5.40,"USDA-adaptado"),
    ("Ovo mexido com manteiga","Ovos",70.5,149,623,9.7,11.5,310,1.6,0,0.8,56,12,0.04,172,1.60,155,138,0.07,1.05,200,200,200,0.07,0.37,0.11,0.06,0,1.17,0.96,1.28,"USDA-adaptado"),
    ("Omelete queijo","Ovos",65.0,196,820,12.5,16.0,360,0.8,0,1.2,145,14,0.03,215,1.92,440,150,0.06,1.55,215,215,215,0.06,0.42,0.12,0.07,0,None,1.00,1.40,"TBCA-USP"),
    ("Albumina em pó","Ovos",8.0,382,1598,80.0,0.5,0,5.0,0,3.5,55,34,0.04,106,0.53,845,900,0.10,0.22,0,0,0,0.03,1.39,0.04,0.65,0,None,0,0.23,"USDA-adaptado"),
    # ─── OLEAGINOSAS E SEMENTES ADICIONAIS ───────────────────────────────────
    ("Castanha de caju torrada sem sal","Oleaginosas",2.0,574,2402,15.3,46.4,0,32.7,3.3,3.4,45,292,1.66,593,6.68,10,565,2.22,5.78,0,0,0,0.30,0.24,0.38,1.52,0.6,None,5.31,0,"USDA-adaptado"),
    ("Avelã","Oleaginosas",5.3,628,2628,15.0,61.0,0,16.7,9.7,2.3,114,163,6.18,290,4.70,0,680,1.73,2.45,1,1,1,0.64,0.11,0.56,1.83,6.3,None,15.28,0,"USDA-adaptado"),
    ("Pinhão cozido","Oleaginosas",46.0,200,837,4.0,9.5,0,38.6,0,0.6,8,68,3.30,66,2.50,3,180,0.36,1.12,0,0,0,0.26,0.07,0.10,1.25,0,None,1.20,0,"TBCA-USP"),
    ("Semente de abóbora","Oleaginosas",7.0,541,2264,24.5,45.8,0,17.8,6.0,4.6,52,550,4.54,1174,8.82,7,807,1.34,7.99,1,1,0,0.27,0.32,0.29,4.99,1.9,None,35.10,0,"USDA-adaptado"),
    ("Semente de girassol","Oleaginosas",4.7,584,2444,20.8,51.5,0,20.0,8.6,3.0,78,325,1.95,660,5.25,9,645,1.80,5.00,1,1,0,1.48,0.36,1.35,8.33,1.4,None,35.17,0,"USDA-adaptado"),
    ("Semente de cânhamo","Oleaginosas",6.6,553,2314,31.6,48.7,0,8.7,4.0,4.8,70,700,7.60,1650,7.95,5,859,1.60,9.90,0,0,0,0.40,0.10,0.12,1.40,0.5,None,1.05,0,"USDA-adaptado"),
    ("Semente de papoula","Oleaginosas",6.8,525,2197,17.9,41.6,0,28.1,19.5,5.7,1438,347,6.71,870,9.76,26,719,1.63,7.90,0,0,0,0.85,0.10,0.25,0.90,1.0,None,0.10,0,"USDA-adaptado"),
    # ─── INGREDIENTES FUNCIONAIS ADICIONAIS ──────────────────────────────────
    ("Spirulina em pó","Aditivos e ingredientes",4.7,290,1214,57.5,7.7,0,23.9,3.6,7.7,120,195,1.90,118,28.50,1048,1363,6.10,2.00,23,23,0,2.38,3.67,0.36,12.82,10.1,None,5.00,0,"USDA-adaptado"),
    ("Clorela em pó","Aditivos e ingredientes",4.0,380,1590,58.4,9.3,0,23.2,0,5.1,221,315,0.30,895,130.0,22,1368,6.70,2.70,0,0,0,1.70,4.00,1.30,20.60,15.6,None,5.00,0,"USDA-adaptado"),
    ("Colágeno tipo II","Carnes e derivados",8.0,360,1507,90.0,0.3,0,0.2,0,1.5,150,5,0.01,50,0.05,80,10,0.03,0.05,0,0,0,0,0.02,0,0,0,None,0,0,"USDA-adaptado"),
    ("HMB (beta-hidroxibutirato)","Aditivos e ingredientes",0,0,0,0,0,0,100,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Creatina monoidratada","Aditivos e ingredientes",0,0,0,87.0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("L-glutamina","Aditivos e ingredientes",0,0,0,100.0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("L-leucina","Aditivos e ingredientes",0,0,0,100.0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("BCAA (Leucina 2:1:1)","Aditivos e ingredientes",0,0,0,100.0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Cafeína anidra","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Extrato de chá verde (EGCG)","Aditivos e ingredientes",5.0,220,921,10.0,1.0,0,45.0,5.0,2.0,50,20,0.50,80,4.00,30,250,0.30,0.50,5,5,3,0.08,0.20,0.10,1.00,0,None,0.50,0,"USDA-adaptado"),
    ("Resveratrol","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Quercetina","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Curcumina (extrato de cúrcuma)","Aditivos e ingredientes",5.0,380,1590,8.0,4.5,0,68.0,10.0,5.0,50,60,2.50,120,10.00,20,800,0.50,1.20,0,0,0,0.05,0.10,0.10,2.00,10.0,None,1.50,0,"USDA-adaptado"),
    ("Piperina (extrato de pimenta)","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Extrato de própolis","Aditivos e ingredientes",30.0,230,963,4.0,14.0,0,40.0,0,5.0,10,10,0.10,30,0.50,20,100,0.10,0.50,0,0,0,0.03,0.05,0.04,0.30,0,None,0.20,0,"USDA-adaptado"),
    ("Probiótico Lactobacillus acidophilus (pó)","Aditivos e ingredientes",5.0,50,209,30.0,2.0,0,5.0,0,5.0,100,20,0.05,80,0.30,200,100,0.05,0.30,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Probiótico Bifidobacterium (pó)","Aditivos e ingredientes",5.0,50,209,30.0,2.0,0,5.0,0,5.0,100,20,0.05,80,0.30,200,100,0.05,0.30,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Beta-glucana de aveia","Aditivos e ingredientes",8.0,150,628,3.0,0.5,0,85.0,85.0,0,30,20,0.10,45,0.50,5,80,0.05,0.20,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Psyllium husk","Aditivos e ingredientes",11.0,217,908,1.9,0.6,0,85.0,85.0,0,56,48,0.84,55,18.18,4,260,0.36,1.70,0,0,0,0.01,0.14,0.20,1.60,0,None,0.60,0,"USDA-adaptado"),
    # ─── ÓLEOS E GORDURAS ESPECIALIZADOS ─────────────────────────────────────
    ("Óleo de abacate","Óleos e gorduras",0,884,3699,0,100,0,0,0,0,0,0,None,0,0,0,0,None,None,0,0,0,0,0,0,0,0,None,14.32,0,"USDA-adaptado"),
    ("Óleo de macadâmia","Óleos e gorduras",0,884,3699,0,100,0,0,0,0,0,0,None,0,0,0,0,None,None,0,0,0,0,0,0,0,0,None,1.50,0,"USDA-adaptado"),
    ("Óleo de sésamo","Óleos e gorduras",0,884,3699,0,100,0,0,0,0,0,0,None,0,0,0,0,None,None,0,0,0,0,0,0,0,0,None,1.40,0,"USDA-adaptado"),
    ("Óleo de cártamo","Óleos e gorduras",0,884,3699,0,100,0,0,0,0,0,0,None,0,0,0,0,None,None,0,0,0,0,0,0,0,0,None,34.10,0,"USDA-adaptado"),
    ("Banha de porco","Óleos e gorduras",0,898,3759,0,99.5,95,0,0,0,0,0,None,0,0.06,0,0,None,None,0,0,0,0,0,0,0,0,None,1.06,0,"USDA-adaptado"),
    ("Gordura de palma fracionada (estearina)","Óleos e gorduras",0,884,3699,0,100,0,0,0,0,0,0,None,0,0,0,0,None,None,0,0,0,0,0,0,0,0,None,2.80,0,"USDA-adaptado"),
    ("Gordura de palma fracionada (oleína)","Óleos e gorduras",0,884,3699,0,100,0,0,0,0,0,0,None,0,0,0,0,None,None,0,0,0,0,0,0,0,0,None,13.00,0,"USDA-adaptado"),
    # ─── LATICÍNIOS P&D — INSUMOS INDUSTRIAIS ─────────────────────────────────
    ("Cloreto de magnésio (nigari)","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,118000,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Sulfato de cálcio (gesso alimentar)","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,23300,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Ácido fosfórico alimentar","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Fosfato monocálcico","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,1600,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Fosfato dissódico","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,14880,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Citrato de sódio","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,26500,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Ácido sórbico","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Riboflavina (vitamina B2)","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1000000,0,0,0,None,0,0,"USDA-adaptado"),
    ("Niacina (vitamina B3)","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1000000,0,None,0,0,"USDA-adaptado"),
    ("Piridoxina (vitamina B6)","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1000000,0,0,None,0,0,"USDA-adaptado"),
    ("Ácido fólico","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Biotina (vitamina B7)","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Tiamina (vitamina B1)","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1000000,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Sulfato de magnésio","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,99000,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Glucono delta-lactona (GDL)","Aditivos e ingredientes",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Transglutaminase (TG)","Aditivos e ingredientes",5.0,100,419,18.0,0.5,0,10.0,0,5.0,30,10,0.05,60,0.20,150,80,0.05,0.20,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Proteína do soro de leite desnaturada","Leite e derivados",4.5,365,1527,75.0,2.0,15,12.0,0,5.0,850,80,0.04,680,0.60,280,580,0.07,4.20,0,0,0,0.22,0.55,0.24,0.88,0,None,0.30,2.30,"USDA-adaptado"),
    ("alfa-lactoalbumina bovina","Leite e derivados",5.0,380,1590,92.0,1.5,15,0,0,4.5,500,50,0.02,400,0.40,100,250,0.04,3.50,0,0,0,0.10,0.35,0.15,0.60,0,None,0.20,1.50,"USDA-adaptado"),
    ("beta-lactoglobulina bovina","Leite e derivados",5.0,378,1582,90.0,2.0,12,0,0,4.0,450,45,0.02,380,0.38,90,240,0.04,3.30,0,0,0,0.09,0.32,0.14,0.55,0,None,0.18,1.40,"USDA-adaptado"),
    ("Concentrado de proteína de batata","Outras",8.0,350,1465,78.0,0.5,0,10.0,2.5,3.5,25,80,0.30,600,3.50,40,1200,0.40,0.80,0,0,0,0.12,0.12,0.40,2.50,20.0,None,0.10,0,"USDA-adaptado"),
    ("Farinha de grão-de-bico","Leguminosas",10.0,387,1619,22.4,6.7,0,57.8,10.8,3.1,105,166,1.81,318,4.86,64,846,0.96,2.81,27,27,16,0.49,0.13,0.43,1.54,3.6,None,0.82,0,"USDA-adaptado"),
    ("Farinha de lentilha vermelha","Leguminosas",9.0,352,1473,25.0,1.5,0,60.0,10.5,3.5,50,110,1.30,440,7.00,5,950,0.72,3.10,2,2,2,0.35,0.20,0.52,2.50,1.5,None,0.48,0,"USDA-adaptado"),
    ("Farinha de soja integral","Leguminosas",8.0,388,1623,34.5,20.6,0,33.7,9.6,5.4,206,280,3.01,632,9.24,13,2255,1.18,4.15,8,8,4,1.07,0.34,0.58,2.72,0,None,1.77,0,"USDA-adaptado"),
    # ─── SUCOS E POLPAS INDUSTRIAIS ───────────────────────────────────────────
    ("Polpa de morango congelada","Frutas",91.0,29,121,0.6,0.3,0,6.6,1.5,0.4,14,12,0.36,22,0.37,1,138,0.05,0.12,0,0,0,0.02,0.02,0.04,0.35,47.0,None,0.26,0,"TBCA-USP"),
    ("Polpa de manga congelada","Frutas",83.0,59,247,0.6,0.4,0,14.5,1.5,0.4,12,14,0.06,12,0.16,3,155,0.08,0.07,350,350,210,0.06,0.04,0.12,0.55,26.0,None,1.00,0,"TBCA-USP"),
    ("Polpa de abacaxi congelada","Frutas",86.0,47,197,0.5,0.2,0,11.7,1.0,0.3,16,16,0.90,8,0.25,1,140,0.07,0.07,2,2,1,0.07,0.04,0.10,0.46,15.0,None,0.02,0,"TBCA-USP"),
    ("Polpa de maracujá congelada","Frutas",79.0,66,276,2.0,1.3,0,14.5,1.5,1.2,14,30,0.10,70,1.60,30,350,0.09,0.12,65,65,39,0.00,0.21,0.11,1.50,28.0,None,0.02,0,"TBCA-USP"),
    ("Polpa de caju congelada","Frutas",85.5,39,163,1.0,0.3,0,9.5,1.5,0.5,5,20,0.15,35,0.45,3,185,0.28,0.30,2,2,1,0.01,0.03,0.18,0.58,175.0,None,0.30,0,"TBCA-USP"),
    ("Polpa de acerola congelada","Frutas",92.0,32,134,0.8,0.2,0,7.0,1.4,0.5,11,17,0.08,16,0.21,6,140,0.16,0.09,35,35,21,0.02,0.05,0.03,0.38,900.0,None,0.10,0,"TBCA-USP"),
    ("Suco de laranja concentrado","Frutas",48.0,179,749,2.5,0.4,0,41.8,0.8,0.9,28,33,0.09,51,0.65,2,560,0.16,0.11,46,46,28,0.26,0.09,0.17,0.92,129.0,None,0.40,0,"USDA-adaptado"),
    ("Suco de tomate","Hortaliças",93.7,17,71,0.8,0.1,0,4.2,0.4,0.6,10,11,0.11,19,0.44,269,217,0.07,0.20,55,55,33,0.05,0.03,0.10,0.75,18.3,None,0.55,0,"USDA-adaptado"),
    # ─── SNACKS E PRODUTOS DE CONVENIÊNCIA ───────────────────────────────────
    ("Chips de batata","Cereais e derivados",2.0,536,2243,7.0,34.6,0,52.9,4.8,2.8,22,52,0.40,159,1.56,524,1642,0.47,1.00,0,0,0,0.20,0.09,0.30,3.33,31.1,None,2.00,0,"USDA-adaptado"),
    ("Pipoca salgada","Cereais e derivados",4.0,459,1921,13.0,22.0,0,59.0,12.0,2.5,7,136,1.00,384,3.19,744,329,0.40,3.22,11,11,6,0.12,0.08,0.24,1.72,0,None,0.42,0,"USDA-adaptado"),
    ("Barra de cereal granola","Cereais e derivados",8.0,415,1737,7.9,16.2,0,64.8,4.5,2.8,68,53,1.60,195,2.85,140,225,0.33,2.00,0,0,0,0.24,0.07,0.11,2.40,0.5,None,3.50,0,"TBCA-USP"),
    ("Biscoito de arroz","Cereais e derivados",5.5,387,1619,8.9,2.8,0,81.7,2.7,0.9,15,45,1.50,200,0.35,110,90,0.15,1.00,0,0,0,0.05,0.03,0.15,2.00,0,None,0.05,0,"USDA-adaptado"),
    ("Barrita proteica (barra proteica)","Outros",10.0,385,1611,25.0,13.5,15,45.0,3.5,2.0,350,60,0.50,350,3.50,250,350,0.35,3.50,0,0,0,0.50,0.65,0.50,5.00,30.0,2.50,3.50,1.50,"USDA-adaptado"),
    # ─── VEGETAIS SECOS E DESIDRATADOS ────────────────────────────────────────
    ("Tomate seco","Hortaliças",14.6,258,1080,14.1,3.0,0,55.8,12.3,10.0,110,194,1.85,356,9.09,2095,3427,1.42,1.99,874,874,524,0.53,0.49,0.72,9.05,39.2,None,7.30,0,"USDA-adaptado"),
    ("Damasco seco","Frutas",30.9,241,1009,3.4,0.5,0,62.6,7.3,2.3,55,32,0.24,71,2.66,10,1162,0.34,0.39,180,180,108,0.02,0.07,0.14,2.58,1.0,None,4.33,0,"USDA-adaptado"),
    ("Uva passa","Frutas",15.0,299,1251,3.1,0.5,0,79.2,3.7,1.7,50,32,0.30,101,1.88,11,749,0.32,0.22,0,0,0,0.11,0.13,0.17,1.14,2.3,None,0.50,0,"USDA-adaptado"),
    ("Ameixa seca","Frutas",31.8,240,1004,2.2,0.4,0,63.9,7.1,2.0,43,41,0.30,69,0.93,2,732,0.28,0.44,39,39,23,0.05,0.19,0.21,1.88,0.6,None,0.43,0,"USDA-adaptado"),
    ("Figo seco","Frutas",30.0,249,1042,3.3,0.9,0,63.9,9.8,2.0,162,68,0.51,67,2.03,10,680,0.29,0.55,9,9,5,0.09,0.08,0.11,0.62,1.2,None,0.35,0,"USDA-adaptado"),
    ("Tâmara seca","Frutas",21.3,282,1180,2.5,0.4,0,75.0,8.0,1.5,39,54,0.30,62,1.02,2,696,0.24,0.29,7,7,4,0.05,0.07,0.19,1.61,0.4,None,0.05,0,"USDA-adaptado"),
    ("Coco desidratado ralado","Frutas",2.7,592,2477,5.7,56.4,0,27.6,18.4,1.9,21,52,2.43,187,3.93,32,578,0.68,1.72,0,0,0,0.07,0.03,0.07,0.86,1.5,None,0.25,0,"USDA-adaptado"),
    # ─── OUTROS ALIMENTOS P&D ─────────────────────────────────────────────────
    ("Leite de soja fermentado","Leguminosas",87.0,48,201,3.8,2.0,0,4.0,0.6,0.8,100,23,0.35,50,0.80,65,145,0.22,0.28,4,4,2,0.08,0.18,0.05,0.18,0.5,1.20,0.20,0.80,"USDA-adaptado"),
    ("Proteína de leite hidrolisada","Leite e derivados",5.5,360,1507,80.0,2.5,18,5.0,0,5.0,900,80,0.04,720,0.62,310,560,0.07,4.30,0,0,0,0.21,0.52,0.23,0.88,0,None,0.30,2.40,"USDA-adaptado"),
    ("Gordura anidra de manteiga fracionada","Óleos e gorduras",0.2,895,3746,0.1,99.5,256,0.1,0,0,22,1,None,4,0.01,4,8,None,None,820,820,820,0,0.01,0,0,0,None,2.45,0.25,"USDA-adaptado"),
    ("Mix de proteínas (whey+caseína)","Leite e derivados",4.5,375,1569,85.0,4.0,22,2.0,0,3.5,700,90,0.05,680,0.85,180,565,0.09,4.60,0,0,0,0.29,0.62,0.29,1.15,0,None,0.48,2.65,"USDA-adaptado"),
    ("Soro de leite concentrado (WPC 35%)","Leite e derivados",4.0,355,1485,35.0,3.5,20,55.5,0,7.5,800,90,0.04,620,0.45,350,750,0.08,4.00,0,0,0,0.25,0.80,0.38,0.75,2.5,None,0.30,1.80,"USDA-adaptado"),
    ("Extrato malte líquido","Cereais e derivados",20.0,316,1322,6.0,0.5,0,73.0,1.5,1.5,30,28,0.50,180,2.00,25,300,0.15,1.10,0,0,0,0.28,0.20,0.30,4.80,0,None,0.25,0,"USDA-adaptado"),
    ("Extrato malte em pó","Cereais e derivados",3.0,368,1540,7.0,0.6,0,87.0,2.0,1.8,35,32,0.60,210,2.30,28,350,0.18,1.30,0,0,0,0.32,0.23,0.35,5.50,0,None,0.30,0,"USDA-adaptado"),
    ("Leite de arroz","Bebidas",88.0,47,197,0.3,1.0,0,9.2,0.3,0.3,118,12,0.27,90,0.24,52,80,0.05,0.28,48,48,14,0.03,0.12,0.06,0.40,0,1.00,1.40,0,"USDA-adaptado"),
    ("Leite de caju","Bebidas",92.0,24,100,0.5,0.5,0,4.8,0.1,0.2,188,6,0.05,20,0.15,40,40,0.04,0.10,45,45,13,0,0.03,0,0,0,0.80,3.40,0,"USDA-adaptado"),
    ("Bebida de coco fermentada","Bebidas",90.0,40,167,0.5,2.0,0,6.0,0.5,0.3,12,15,0.10,15,0.20,40,160,0.05,0.10,0,0,0,0.02,0.02,0.02,0.10,1.5,None,0.05,0,"USDA-adaptado"),
    ("Creme vegetal de aveia","Bebidas",75.0,65,272,1.5,3.5,0,8.5,1.0,0.5,140,12,0.20,95,0.30,55,100,0.04,0.30,45,45,14,0.06,0.15,0.03,0.08,0,0.75,1.00,0,"USDA-adaptado"),
    ("Iogurte de coco","Leite e derivados",78.0,110,460,0.5,8.5,0,10.5,0.5,0.8,12,18,0.50,12,0.40,15,140,0.12,0.10,0,0,0,0.02,0.01,0.02,0.15,1.5,None,0.08,0,"USDA-adaptado"),
    ("Queijo vegano de castanha","Outros",55.0,210,879,6.5,16.0,0,14.5,3.5,3.5,80,80,1.00,190,2.20,480,280,0.70,1.50,0,0,0,0.10,0.15,0.12,1.50,0,None,3.50,0,"USDA-adaptado"),
    ("Margarina vegana","Óleos e gorduras",16.0,717,3000,0.1,79.0,0,0.5,0,1.0,8,1,0.01,8,0.05,760,28,0.01,0.06,300,300,0,0,0.01,0,0,0,None,15.00,0,"USDA-adaptado"),
    ("Creme azedo (crème fraîche)","Leite e derivados",72.0,198,829,2.4,19.0,65,3.5,0,0.6,100,9,0.01,72,0.04,45,140,0.01,0.30,188,188,56,0.03,0.13,0.03,0.06,0.4,None,0.75,0.22,"USDA-adaptado"),
    ("Leitelho em pó (buttermilk)","Leite e derivados",3.5,375,1569,34.3,5.8,25,48.9,0,7.8,1200,100,0.02,950,0.35,430,1650,0.08,4.30,8,8,2,0.35,1.48,0.43,0.86,5.5,None,0.08,3.45,"USDA-adaptado"),
    ("Leitelho líquido","Leite e derivados",90.1,40,167,3.3,0.9,4,4.8,0,0.7,116,11,0.01,89,0.05,105,151,0.01,0.42,12,12,4,0.04,0.16,0.04,0.09,0.8,None,0.03,0.44,"USDA-adaptado"),
    ("Proteína de ervilha texturizada","Leguminosas",8.0,360,1507,80.0,5.0,0,5.0,2.0,2.5,60,130,0.60,720,9.60,700,310,0.38,3.00,0,0,0,0.55,0.22,0.33,1.65,0,None,0.55,0,"USDA-adaptado"),
    ("Fibra de chicória (inulina chicória)","Aditivos e ingredientes",4.0,175,733,0.1,0,0,88.0,88.0,0,30,12,0.04,24,0.15,5,78,0.04,0.08,0,0,0,0,0,0,0,0,None,0,0,"USDA-adaptado"),
    ("Farinha de linhaça","Oleaginosas",6.9,450,1883,28.8,31.5,0,28.8,27.3,3.4,255,392,2.48,642,5.73,30,813,1.22,4.34,0,0,0,1.64,0.16,0.47,3.08,0.6,None,31.51,0,"USDA-adaptado"),
    ("Farinha de chia","Oleaginosas",6.0,450,1883,20.0,28.5,0,39.0,32.0,4.0,580,310,2.50,785,7.00,14,380,0.85,4.20,0,0,0,0.57,0.16,0.18,8.10,1.5,None,0.46,0,"USDA-adaptado"),
]


def baixar_taco_excel(destino: str, timeout: float = 20.0) -> bool:
    """Tenta baixar a planilha TACO 4ª edição do servidor UNICAMP.

    Tem tempo limite (sem ele uma conexão travada segurava a inicialização do
    app indefinidamente) e grava em arquivo temporário, para que um download
    interrompido não deixe uma planilha corrompida no lugar.
    """
    tmp = destino + ".parcial"
    try:
        import urllib.request
        print(f"  → Baixando TACO de {URL_TACO} ...")
        with urllib.request.urlopen(URL_TACO, timeout=timeout) as resp, open(tmp, "wb") as f:
            while True:
                bloco = resp.read(64 * 1024)
                if not bloco:
                    break
                f.write(bloco)
        os.replace(tmp, destino)
        size_kb = os.path.getsize(destino) // 1024
        print(f"  → Download concluído ({size_kb} KB).")
        return True
    except Exception as e:
        print(f"  ⚠️  Não foi possível baixar TACO automaticamente: {e}")
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except OSError:
            pass
        return False


def criar_taco_db():
    """Cria e popula o banco de dados TACO.

    O banco é montado em um arquivo temporário e só então renomeado para
    taco.db: se o processo cair no meio, não sobra um banco pela metade.
    """
    print(f"→ Criando banco TACO em: {TACO_DB}")
    tmp_db = TACO_DB + ".novo"
    for resto in (tmp_db, tmp_db + "-wal", tmp_db + "-shm"):
        if os.path.exists(resto):
            os.remove(resto)

    conn = sqlite3.connect(tmp_db)
    conn.executescript(SQL_TACO_DB)

    # 1) Tentar arquivo local; se não existir, baixar
    if not os.path.exists(TACO_XLS):
        baixar_taco_excel(TACO_XLS)

    # 2) Importar do Excel (local ou recém-baixado)
    dados_carregados = False
    if os.path.exists(TACO_XLS):
        try:
            dados_carregados = importar_taco_excel(conn, TACO_XLS)
        except Exception as e:
            print(f"  ⚠️  Falha ao importar Excel TACO: {e}")
            print("  → Usando dados embutidos como fallback.")

    # 3) Fallback: base embutida
    if not dados_carregados:
        print("  → Usando base embutida como fallback.")
        importar_taco_embutida(conn)

    # 4) Sempre inserir suplemento (ingredientes especializados P&D)
    importar_suplemento(conn)

    conn.commit()
    count = conn.execute("SELECT COUNT(*) FROM alimentos_taco").fetchone()[0]
    conn.close()
    os.replace(tmp_db, TACO_DB)
    print(f"  ✅ {count} alimentos carregados na TACO.")


def _taco_float(val) -> float | None:
    """Converte célula TACO para float. 'Tr' (traço) → 0.0; vazio/NA → None."""
    if val is None:
        return None
    s = str(val).strip()
    if s in ('', 'nan', 'NaN', 'NA', 'NA ', 'None'):
        return None
    if s.lower() in ('tr', 'tr.', '*'):
        return 0.0
    s = s.replace(',', '.').replace(' ', '')
    try:
        return float(s)
    except ValueError:
        return None


def importar_taco_excel(conn: sqlite3.Connection, caminho: str) -> bool:
    """
    Importa planilha TACO 4ª edição (formato oficial UNICAMP/NEPA) para SQLite.
    O arquivo tem cabeçalho mesclado em 3 linhas; usa mapeamento por posição.
    """
    try:
        df = pd.read_excel(caminho, header=None, dtype=str)
    except Exception as e:
        raise ValueError(f"Erro ao ler Excel: {e}")

    all_cols = ['nome_alimento', 'categoria'] + list(COL_NUTRIENTES.keys())
    placeholders = ', '.join(['?' for _ in all_cols])
    col_names    = ', '.join(all_cols)

    # Textos que identificam linhas de cabeçalho/legenda repetidas — não são categorias
    TEXTOS_CABECALHO = {
        'número do', 'alimento', 'descrição dos alimentos', 'umidade',
        'energia', 'proteína', 'lipídeos', 'carboidrato', 'legenda',
        '(%)', '(g)', '(mg)', '(kcal)', '(kj)', 'vitamina',
    }

    conn.execute("DELETE FROM alimentos_taco")
    inseridos = 0
    categoria = ""

    for idx, row in df.iterrows():
        if idx < 3:          # pular as 3 linhas de cabeçalho
            continue

        col0_raw  = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        nome      = str(row.iloc[COL_NOME]).strip() if pd.notna(row.iloc[COL_NOME]) else ""

        # Linha de dado: col 0 deve ser numérica (ID do alimento)
        try:
            float(col0_raw.replace(',', '.'))
            eh_dado = True
        except ValueError:
            eh_dado = False

        if not eh_dado:
            # Pode ser linha de categoria — ignorar se for texto de cabeçalho
            if nome.lower() in ('nan', ''):
                if col0_raw and col0_raw.lower() not in TEXTOS_CABECALHO:
                    # Verificar que parece uma categoria real (texto sem números)
                    if not re.search(r'\d', col0_raw) and len(col0_raw) > 3:
                        categoria = col0_raw
            continue

        if not nome or nome.lower() in ('nan', ''):
            continue

        vals = [nome, categoria]
        for nut in COL_NUTRIENTES:
            col_idx = COL_NUTRIENTES[nut]
            vals.append(_taco_float(row.iloc[col_idx]) if col_idx < len(row) else None)

        conn.execute(f"INSERT INTO alimentos_taco ({col_names}) VALUES ({placeholders})", vals)
        inseridos += 1

    print(f"  → {inseridos} alimentos importados do Excel TACO.")
    return inseridos > 0


def importar_taco_embutida(conn: sqlite3.Connection):
    """Insere os dados TACO embutidos no código."""
    conn.execute("DELETE FROM alimentos_taco")
    cols = [col for col, _ in COLUNAS_TACO]
    placeholders = ", ".join(["?" for _ in cols])
    col_names    = ", ".join(cols)
    for registro in TACO_EMBUTIDA:
        conn.execute(f"INSERT INTO alimentos_taco ({col_names}) VALUES ({placeholders})", registro)
    print(f"  → {len(TACO_EMBUTIDA)} alimentos da base embutida inseridos.")


def completar_suplemento() -> int:
    """
    Acrescenta a um taco.db JÁ EXISTENTE os itens do suplemento que ainda não
    estão nele (ex.: itens novos após uma atualização do app). Não altera nem
    remove nada. Retorna quantos itens foram incluídos.
    """
    if not os.path.exists(TACO_DB):
        return 0
    import unicodedata

    def _chave(nome: str) -> str:
        return unicodedata.normalize("NFKD", nome).encode("ascii", "ignore").decode().lower().strip()

    cols = [col for col, _ in COLUNAS_TACO] + ['fonte']
    sql = (f"INSERT INTO alimentos_taco ({', '.join(cols)}) "
           f"VALUES ({', '.join('?' for _ in cols)})")
    conn = sqlite3.connect(TACO_DB)
    try:
        existentes = {_chave(r[0]) for r in conn.execute("SELECT nome_alimento FROM alimentos_taco")}
        novos = [r for r in TACO_SUPLEMENTO + TACO_SUPLEMENTO_2 if _chave(r[0]) not in existentes]
        if novos:
            conn.executemany(sql, novos)
            conn.commit()
        return len(novos)
    finally:
        conn.close()


def importar_suplemento(conn: sqlite3.Connection):
    """Insere suplemento de ingredientes especializados (TBCA-USP / USDA-adaptado)."""
    cols = [col for col, _ in COLUNAS_TACO] + ['fonte']
    placeholders = ", ".join(["?" for _ in cols])
    col_names    = ", ".join(cols)
    count = 0
    todos = TACO_SUPLEMENTO + TACO_SUPLEMENTO_2
    for registro in todos:
        try:
            conn.execute(
                f"INSERT INTO alimentos_taco ({col_names}) VALUES ({placeholders})",
                registro
            )
            count += 1
        except Exception as e:
            print(f"  ⚠️  Erro ao inserir '{registro[0]}': {e}")
    print(f"  → {count} alimentos suplementares inseridos.")


_DENSIDADES_INICIAIS = [
    ("Leite integral UHT",  1.030),
    ("Leite desnatado",     1.035),
    ("Creme de leite 35%",  1.012),
    ("Iogurte integral",    1.050),
    ("Soro whey",           1.025),
    ("Leite condensado",    1.300),
    ("Óleo de soja",        0.915),
    ("Óleo de palma",       0.910),
    ("Azeite",              0.911),
    ("Óleo de girassol",    0.920),
    ("Água",                1.000),
    ("Suco de laranja",     1.045),
    ("Vinagre",             1.006),
    ("Mel",                 1.420),
    ("Glicerina",           1.261),
    ("Álcool etílico",      0.789),
    ("Xarope de glicose",   1.380),
    ("Extrato de malte",    1.200),
]


def criar_app_db():
    """Cria o banco de dados da aplicação (receitas, fornecedores, densidades)."""
    print(f"→ Criando banco da aplicação em: {APP_DB}")
    conn = sqlite3.connect(APP_DB)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    conn.executescript(SQL_APP_DB)
    conn.executemany(
        "INSERT OR IGNORE INTO densidades_ingredientes (nome_ingrediente, densidade_g_ml, fonte) VALUES (?,?,'tabela')",
        _DENSIDADES_INICIAIS,
    )
    conn.commit()
    conn.close()
    print("  ✅ Banco da aplicação criado.")


if __name__ == "__main__":
    print("=" * 55)
    print("  NutriCalc P&D — Inicialização dos Bancos de Dados")
    print("=" * 55)
    criar_taco_db()
    criar_app_db()
    print()
    print("✅ Setup concluído! Execute: streamlit run app.py")
    print("=" * 55)
