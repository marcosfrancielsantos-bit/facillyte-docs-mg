import base64
import io
import json
import os
import re
import sqlite3
import subprocess
import tempfile
import time
import zipfile
from datetime import datetime

import docx
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Inches, Mm, Pt, RGBColor
import pandas as pd
import streamlit as st

try:
    from google import genai
    from google.genai import types
    GENAI_AVAILABLE = True
except ImportError:
    GENAI_AVAILABLE = False

# ==============================================================================
# CONFIGURAÇÃO E IDENTIDADE VISUAL FACILLYTE ENGENHARIA & TOPOGRAFIA
# ==============================================================================
st.set_page_config(
    page_title="Facillyte Docs — Funcionário Virtual Cartorial (MG)",
    page_icon="📐",
    layout="wide",
    initial_sidebar_state="expanded",
)

DB_PATH = "facillyte_processos.db"


def find_logo(filename: str) -> str:
    for candidate in (filename, os.path.join("assets", filename)):
        if os.path.exists(candidate):
            return candidate
    return filename


def img_to_b64(path: str) -> str:
    if not os.path.exists(path):
        return ""
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


LOGO_ENG = find_logo("logo_facillyte_engenharia.png")
LOGO_TOPO = find_logo("logo_facillyte_topografia.png")

HEX_PRIMARY = "0B2457"      # Azul-Marinho Oficial Facillyte
HEX_ACCENT = "0A8A1C"       # Verde Topográfico Oficial Facillyte
HEX_SLATE = "1E293B"
HEX_MUTED = "475569"
HEX_TINT = "F1F5F9"
HEX_BORDER = "CBD5E1"

RGB_PRIMARY = RGBColor(0x0B, 0x24, 0x57)
RGB_ACCENT = RGBColor(0x0A, 0x8A, 0x1C)
RGB_SLATE = RGBColor(0x1E, 0x29, 0x3B)
RGB_MUTED = RGBColor(0x47, 0x55, 0x69)
RGB_WHITE = RGBColor(0xFF, 0xFF, 0xFF)


def injetar_tema_facillyte():
    st.markdown(
        """
        <style>
        .stApp { background-color: #F4F7FA; color: #1E293B; }
        section[data-testid="stSidebar"] {
            background: linear-gradient(180deg, #0B2457 0%, #0F2F6E 100%);
            color: #FFFFFF;
        }
        section[data-testid="stSidebar"] * { color: #F8FAFC !important; }
        section[data-testid="stSidebar"] .stTextInput input {
            background-color: rgba(255,255,255,0.12) !important;
            color: #FFFFFF !important;
            border: 1px solid rgba(255,255,255,0.3) !important;
            border-radius: 8px !important;
        }
        section[data-testid="stSidebar"] button {
            border-radius: 8px !important;
            border: 1px solid rgba(255,255,255,0.22) !important;
            background-color: rgba(255,255,255,0.08) !important;
            transition: all 0.2s ease;
        }
        section[data-testid="stSidebar"] button:hover {
            background-color: #0A8A1C !important;
            border-color: #0A8A1C !important;
        }
        section[data-testid="stSidebar"] button[kind="primary"] {
            background-color: #0A8A1C !important;
            border-color: #0A8A1C !important;
            font-weight: 600 !important;
        }
        .fac-banner {
            background: #FFFFFF;
            border-top: 5px solid #0B2457;
            border-bottom: 3px solid #0A8A1C;
            border-radius: 12px;
            padding: 14px 22px;
            margin-bottom: 18px;
            box-shadow: 0 3px 12px rgba(11, 36, 87, 0.07);
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 16px;
        }
        .fac-banner-center { text-align: center; flex: 1; }
        .fac-banner-title { color: #0B2457; font-size: 1.35rem; font-weight: 800; margin: 0; }
        .fac-banner-sub { color: #0A8A1C; font-size: 0.92rem; font-weight: 600; margin: 2px 0 0 0; }
        .fac-pill {
            display: inline-block; background: #E8F5E9; color: #0A8A1C;
            border: 1px solid #0A8A1C; padding: 3px 12px; border-radius: 999px;
            font-size: 0.82rem; font-weight: 700; margin-top: 6px;
        }
        .fac-card {
            background: #FFFFFF; border: 1px solid #DCE4EE; border-left: 5px solid #0B2457;
            border-radius: 10px; padding: 14px 16px; margin-bottom: 10px;
            box-shadow: 0 2px 6px rgba(11, 36, 87, 0.04);
        }
        .fac-card-green { border-left-color: #0A8A1C !important; }
        .fac-card-amber { border-left-color: #D97706 !important; background: #FFFBEB; }
        .fac-card h4 { margin: 0 0 4px 0; color: #0B2457; font-size: 1.03rem; }
        .fac-card p { margin: 0; color: #475569; font-size: 0.88rem; }
        .stButton > button[kind="primary"], .stFormSubmitButton > button[kind="primary"] {
            background: linear-gradient(90deg, #0B2457 0%, #143A82 100%) !important;
            color: #FFFFFF !important; border: none !important; border-radius: 9px !important;
            font-weight: 600 !important; padding: 0.55rem 1.1rem !important;
            box-shadow: 0 2px 8px rgba(11, 36, 87, 0.22) !important;
        }
        .stButton > button[kind="primary"]:hover, .stFormSubmitButton > button[kind="primary"]:hover {
            background: linear-gradient(90deg, #0A8A1C 0%, #0D9E22 100%) !important;
        }
        .stButton > button[kind="secondary"] {
            background: #FFFFFF !important; color: #0B2457 !important;
            border: 1.5px solid #CBD5E1 !important; border-radius: 9px !important; font-weight: 600 !important;
        }
        .stButton > button[kind="secondary"]:hover {
            border-color: #0A8A1C !important; color: #0A8A1C !important;
        }
        .stDownloadButton > button {
            background: linear-gradient(90deg, #0A8A1C 0%, #0E9F24 100%) !important;
            color: #FFFFFF !important; border: none !important; border-radius: 10px !important;
            font-weight: 700 !important; font-size: 1.02rem !important; padding: 0.65rem 1.2rem !important;
        }
        .stTabs [data-baseweb="tab-list"] { gap: 8px; background: transparent; }
        .stTabs [data-baseweb="tab"] {
            background: #FFFFFF; border: 1px solid #CBD5E1; border-radius: 8px 8px 0 0;
            padding: 8px 16px; color: #0B2457; font-weight: 600;
        }
        .stTabs [aria-selected="true"] {
            background: #0B2457 !important; color: #FFFFFF !important; border-color: #0B2457 !important;
        }
        div[data-testid="stMetric"] {
            background: #FFFFFF; border: 1px solid #DCE4EE; border-top: 4px solid #0B2457;
            border-radius: 10px; padding: 12px 16px; box-shadow: 0 2px 6px rgba(11, 36, 87, 0.04);
        }
        .stProgress > div > div > div > div {
            background: linear-gradient(90deg, #0B2457 0%, #0A8A1C 100%) !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_banner(titulo: str, subtitulo: str, badge: str = ""):
    b64_eng = img_to_b64(LOGO_ENG)
    b64_topo = img_to_b64(LOGO_TOPO)
    img_eng = f'<img src="data:image/png;base64,{b64_eng}" style="height:56px;object-fit:contain;">' if b64_eng else '<strong style="color:#0B2457;">FACILLYTE<br>ENGENHARIA</strong>'
    img_topo = f'<img src="data:image/png;base64,{b64_topo}" style="height:60px;object-fit:contain;">' if b64_topo else '<strong style="color:#0A8A1C;">FACILLYTE<br>TOPOGRAFIA</strong>'
    badge_html = f'<div class="fac-pill">{badge}</div>' if badge else ""
    st.markdown(
        f"""
        <div class="fac-banner">
            <div>{img_eng}</div>
            <div class="fac-banner-center">
                <div class="fac-banner-title">{titulo}</div>
                <div class="fac-banner-sub">{subtitulo}</div>
                {badge_html}
            </div>
            <div>{img_topo}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ==============================================================================
# CATÁLOGO DE RITOS (MG) + DOCUMENTOS FÍSICOS EXIGIDOS POR LEGISLAÇÃO (CGJ/MG)
# ==============================================================================
RITOS_MG = {
    "Georreferenciamento (SIGEF/INCRA)": {
        "icone": "🛰️",
        "fundamento": "Art. 176, §§ 3º a 5º, e art. 213 da Lei nº 6.015/1973; Lei nº 10.267/2001; Decretos nº 4.449/2002 e 5.570/2005; arts. 891 (§ 9º) e 1.027+ do Provimento Conjunto nº 93/2020 (CGJ/MG).",
        "docs_proprietario": "RG, CPF, Certidão de Casamento/Nascimento (com regime de bens), Pacto Antenupcial (se houver), Endereço; PJ: Contrato Social; Espólio: Termo de Inventariante + Óbito.",
        "docs_imovel": "Matrícula atualizada, CCIR quitado, CND do ITR, Recibo CAR, Certificação/Planilha SIGEF (.xlsx/.ods/.pdf) e ART quitada.",
        "docs_confrontante": "RG/CPF ou CNPJ, endereço, nº da Matrícula/Posse; se Posseiro: recibo/contrato/CCIR/ITR/CAR autenticado; se Espólio: Inventariante + Óbito; se PJ: Contrato Social.",
        "docs_fisicos_base": [
            "Requerimento original assinado pelo(s) proprietário(s), cônjuge e RT, com firmas reconhecidas (ou assinatura presencial no CRI — arts. 861 e 877 do Prov. Conj. nº 93/2020).",
            "Planta(s) e Memorial(is) Descritivo(s) certificados pelo SIGEF/INCRA impressos em A4, assinados por proprietário, RT e confrontantes (firmas reconhecidas).",
            "Declarações individuais de anuência dos confrontantes com firma reconhecida (ou pedido de notificação via RTD/edital).",
            "ART quitada e assinada (CREA-MG).",
            "CCIR atualizado e quitado + CND do ITR válida (ou 5 últimos comprovantes) + Recibo do CAR.",
            "Declaração de Valor Real e de Mercado assinada com firma reconhecida (art. 10, § 3º, XVI, Lei Estadual nº 15.424/2004 — Tabela 4).",
            "Cópias autenticadas de RG, CPF e Certidão de Casamento/Nascimento atualizada (e Certidão de Registro do Pacto Antenupcial se regime diverso da comunhão parcial).",
        ],
        "exige_confrontantes": True,
        "exige_fmp": False,
        "exige_valor_mercado": True,
    },
    "Retificação Administrativa de Área": {
        "icone": "📏",
        "fundamento": "Art. 213, II, da Lei nº 6.015/1973 e arts. 889 a 900 (art. 891) do Provimento Conjunto nº 93/2020 (CGJ/MG).",
        "docs_proprietario": "RG, CPF, Certidão de Casamento/Nascimento (com regime de bens), Pacto Antenupcial (se houver) e Comprovante de Endereço.",
        "docs_imovel": "Matrícula atualizada, CCIR quitado, CND do ITR, Recibo CAR, Certificação SIGEF (quando exigível) e ART quitada.",
        "docs_confrontante": "RG/CPF ou CNPJ, estado civil, endereço e nº da Matrícula/Posse (se posseiro: prova de posse autenticada).",
        "docs_fisicos_base": [
            "Requerimento original com firma reconhecida (proprietário, cônjuge e RT — arts. 861, 877 e 891 do Prov. Conj. nº 93/2020).",
            "Planta e Memorial(is) Descritivo(s) em SIRGAS 2000 impressos (e arquivo digital .docx quando exigido), com atestados do RT (in loco) e proprietário e firmas reconhecidas.",
            "Declarações individuais de anuência de todos os confrontantes (art. 213, § 10, LRP) com firmas reconhecidas.",
            "ART quitada e assinada + CCIR quitado + CND ITR + Recibo CAR (+ SIGEF se imóvel rural sujeito à certificação).",
            "Declaração de Valor Real e de Mercado (Lei Estadual nº 15.424/2004, Tabela 4) com firma reconhecida.",
            "Cópias autenticadas de documentos pessoais e Certidão de Estado Civil atualizada.",
        ],
        "exige_confrontantes": True,
        "exige_fmp": False,
        "exige_valor_mercado": True,
    },
    "Desmembramento Rural": {
        "icone": "🌾",
        "fundamento": "Arts. 176 e 246 da Lei nº 6.015/1973; art. 65 da Lei nº 4.504/1964 (FMP); art. 991 do Provimento Conjunto nº 93/2020 (CGJ/MG).",
        "docs_proprietario": "RG, CPF, Certidão de Casamento/Nascimento (com regime de bens) e Comprovante de Endereço.",
        "docs_imovel": "Matrícula atualizada, CCIR quitado, CND do ITR, Recibo CAR, Certificações SIGEF das glebas (quando exigível) e ART quitada.",
        "docs_confrontante": "Opcional se não houver alteração do perímetro externo (obrigatório caso cumule com retificação).",
        "docs_fisicos_base": [
            "Requerimento de Desmembramento original assinado pelo(s) proprietário(s), cônjuge e RT com firmas reconhecidas.",
            "Planta geral e Memoriais Descritivos DISTINTOS de cada Gleba e da Área Remanescente (SIRGAS 2000), comprovando que todas respeitam a FMP (art. 65 Lei 4.504/64 e art. 991 Prov. Conj. 93/2020).",
            "ART quitada e assinada + Certificações SIGEF das parcelas (quando exigível).",
            "CCIR quitado + CND do ITR + Recibo do CAR.",
            "Declaração de Valor Real e de Mercado discriminando o valor de cada gleba e do remanescente (Lei Estadual nº 15.424/2004).",
        ],
        "exige_confrontantes": False,
        "exige_fmp": True,
        "exige_valor_mercado": True,
    },
    "Estremação (Condomínio Pro Diviso)": {
        "icone": "🧭",
        "fundamento": "Arts. 1.149 a 1.155 do Provimento Conjunto nº 93/2020 (CGJ/MG) e art. 65 da Lei nº 4.504/1964 (FMP).",
        "docs_proprietario": "RG, CPF, Certidão de Casamento/Nascimento, prova da fração ideal/posse localizada (> 5 anos) e endereço.",
        "docs_imovel": "Matrícula mãe em condomínio, CCIR, CND ITR, CAR, SIGEF (se aplicável), ART quitada e Escritura Pública de Estremação.",
        "docs_confrontante": "RG/CPF, endereço e matrícula dos confrontantes lindeiros (internos do condomínio e externos).",
        "docs_fisicos_base": [
            "Escritura Pública Declaratória de Estremação lavrada em Tabelionato de Notas (arts. 1.149 a 1.155 do Prov. Conj. nº 93/2020).",
            "Planta e Memorial(is) Descritivo(s) da parcela estremada (comprovando parcela e remanescente ≥ FMP do município).",
            "Anuência expressa de todos os confrontantes lindeiros (internos ao condomínio e externos) com firmas reconhecidas.",
            "ART quitada + CCIR quitado + CND do ITR + Recibo CAR + SIGEF (quando exigível).",
            "Declaração de Valor Real e de Mercado (Lei Estadual nº 15.424/2004).",
        ],
        "exige_confrontantes": True,
        "exige_fmp": True,
        "exige_valor_mercado": True,
    },
    "Inserção de Dados / Medidas Atualizadas": {
        "icone": "📐",
        "fundamento": "Art. 213, I e II, da Lei nº 6.015/1973 e art. 891 do Provimento Conjunto nº 93/2020 (CGJ/MG).",
        "docs_proprietario": "RG, CPF, Certidão de Casamento/Nascimento (filiação, profissão e regime de bens) e Comprovante de Endereço.",
        "docs_imovel": "Matrícula/Transcrição atualizada, CCIR, CND ITR, CAR (se rural) ou CTM/IPTU (se urbano) e ART quitada.",
        "docs_confrontante": "RG/CPF, endereço e nº da Matrícula/Transcrição de cada confrontante lindeiro.",
        "docs_fisicos_base": [
            "Requerimento original de Inserção de Medidas com firma reconhecida (proprietário, cônjuge e RT).",
            "Planta e Memorial(is) Descritivo(s) (SIRGAS 2000) com assinaturas e firmas reconhecidas.",
            "Declarações de anuência dos confrontantes + ART quitada.",
            "CCIR, CND ITR e CAR (imóvel rural) ou certidão/CTM municipal (imóvel urbano).",
            "Declaração de Valor Real e de Mercado (Lei Estadual nº 15.424/2004).",
        ],
        "exige_confrontantes": True,
        "exige_fmp": False,
        "exige_valor_mercado": True,
    },
    "Regularização de Posse / Usucapião Extrajudicial (Art. 216-A)": {
        "icone": "📜",
        "fundamento": "Art. 216-A da Lei nº 6.015/1973; Provimento nº 149/2023/CNJ (arts. 401+); arts. 264 e 1.157 do Provimento Conjunto nº 93/2020 (CGJ/MG).",
        "docs_proprietario": "RG, CPF, Certidão de Estado Civil, Recibos/Contratos/Declaração de Posse (originais ou autenticados), contas antigas de luz/água e procuração ao advogado (OAB).",
        "docs_imovel": "Certidão negativa de matrícula (ou matrícula atingida), CCIR, 5 últimos ITR/CND ITR, Recibo CAR, SIGEF/INCRA (se rural), Ata Notarial e ART quitada.",
        "docs_confrontante": "RG/CPF, endereço, matrícula/posse e assinatura na planta/memorial ou declaração de anuência com firma reconhecida.",
        "docs_fisicos_base": [
            "Ata Notarial lavrada em Tabelionato de Notas atestando tempo e cadeia de posse (art. 264 do Prov. Conj. nº 93/2020; mínimo de 2 testemunhas se lote vago/sem edificação — art. 1.157, § 1º).",
            "Requerimento assinado por Advogado (OAB) com procuração específica e outorga conjugal quando exigida (art. 401, VII, Prov. nº 149/2023/CNJ).",
            "Documentos comprobatórios da posse em ORIGINAL ou CÓPIA AUTENTICADA: recibos/contratos de compra e venda, cessão de posse, declaração de posse com firma reconhecida, contas de energia/água e notas fiscais rurais.",
            "Planta e Memorial(is) Descritivo(s) em SIRGAS 2000 assinados pelo RT, requerente e confrontantes (firmas reconhecidas) + ART quitada.",
            "Cadastros rurais obrigatórios em MG: CCIR quitado, CND do ITR (ou 5 últimos recibos de pagamento do ITR), Recibo do CAR e Certificação SIGEF/INCRA (quando exigível). (Se urbano: Certidão de Cadastro Imobiliário Municipal/IPTU).",
            "Certidões negativas dos distribuidores (Justiça Estadual e Federal) da comarca do imóvel e do domicílio do requerente (últimos 30 dias).",
            "Declaração de Valor Real e de Mercado (Lei Estadual nº 15.424/2004).",
        ],
        "exige_confrontantes": True,
        "exige_fmp": False,
        "exige_valor_mercado": True,
    },
    "Averbação Diversa (CAR / CCIR / Construção / Subjetiva)": {
        "icone": "📝",
        "fundamento": "Arts. 167, II, e 246 da Lei nº 6.015/1973; art. 717 do Provimento Conjunto nº 93/2020 (CGJ/MG); Provimento nº 61/2017/CNJ.",
        "docs_proprietario": "RG, CPF, Certidão de Casamento/Nascimento atualizada e Comprovante de Endereço.",
        "docs_imovel": "Matrícula atualizada + documento específico do ato (CAR, CCIR, Habite-se/CND INSS, Certidão Civil, etc.).",
        "docs_confrontante": "Não exigido para averbações simples.",
        "docs_fisicos_base": [
            "Requerimento original assinado pelo proprietário/interessado com firma reconhecida (arts. 861 e 877 do Prov. Conj. nº 93/2020).",
            "Documento oficial comprobatório do ato em original ou cópia autenticada (Recibo CAR, CCIR, Habite-se + CND INSS/SERO, Certidão de Casamento/Óbito atualizada, etc.).",
            "Declaração de Valor (apenas se o ato possuir conteúdo financeiro na Tabela 4 da Lei Estadual nº 15.424/2004).",
        ],
        "exige_confrontantes": False,
        "exige_fmp": False,
        "exige_valor_mercado": False,
    },
}



# ==============================================================================
# CONVERSOR AUTOMÁTICO DE VALOR EM REAIS PARA EXTENSO (PT-BR)
# ==============================================================================
def parse_brl(valor_str) -> float | None:
    if valor_str is None:
        return None
    s = str(valor_str).strip()
    if not s:
        return None
    s = re.sub(r"[^0-9,\.]", "", s)
    if not s:
        return None
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    elif s.count(".") > 1:
        s = s.replace(".", "")
    elif "." in s:
        partes = s.split(".")
        if len(partes[-1]) == 3:
            s = s.replace(".", "")
    try:
        return round(float(s), 2)
    except ValueError:
        return None


def formatar_brl(valor: float) -> str:
    s = f"{valor:,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def _centena_extenso(n: int) -> str:
    unidades = ["", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove"]
    especiais = ["dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis", "dezessete", "dezoito", "dezenove"]
    dezenas = ["", "", "vinte", "trinta", "quarenta", "cinquenta", "sessenta", "setenta", "oitenta", "noventa"]
    centenas = ["", "cento", "duzentos", "trezentos", "quatrocentos", "quinhentos", "seiscentos", "setecentos", "oitocentos", "novecentos"]
    if n == 0:
        return ""
    if n == 100:
        return "cem"
    c, resto = divmod(n, 100)
    d, u = divmod(resto, 10)
    partes = []
    if c > 0:
        partes.append(centenas[c])
    if 10 <= resto <= 19:
        partes.append(especiais[resto - 10])
    else:
        if d > 0:
            partes.append(dezenas[d])
        if u > 0:
            partes.append(unidades[u])
    return " e ".join(partes)


def _inteiro_extenso(n: int) -> str:
    if n == 0:
        return "zero"
    escalas = [
        (1_000_000_000, "bilhão", "bilhões"),
        (1_000_000, "milhão", "milhões"),
        (1_000, "mil", "mil"),
        (1, "", ""),
    ]
    grupos = []
    resto = n
    for div, sing, plur in escalas:
        q, resto = divmod(resto, div)
        if q == 0:
            continue
        if div == 1_000 and q == 1:
            grupos.append((div, q, "mil"))
        else:
            txt = _centena_extenso(q)
            nome = sing if q == 1 else plur
            grupos.append((div, q, f"{txt} {nome}".strip()))
    res = ""
    for i, (div, q, txt) in enumerate(grupos):
        if i == 0:
            res = txt
        else:
            if div == 1 and (q < 100 or q % 100 == 0):
                res += " e " + txt
            elif i == len(grupos) - 1 and div >= 1_000:
                res += " e " + txt
            else:
                res += ", " + txt
    return res


def valor_por_extenso_brl(valor_str) -> str:
    val = parse_brl(valor_str)
    if val is None or val < 0:
        return ""
    reais = int(val)
    centavos = int(round((val - reais) * 100))
    partes = []
    if reais > 0:
        ext_r = _inteiro_extenso(reais)
        if reais >= 1_000_000 and reais % 1_000_000 == 0:
            partes.append(f"{ext_r} de reais")
        else:
            partes.append(f"{ext_r} {'real' if reais == 1 else 'reais'}")
    if centavos > 0:
        ext_c = _centena_extenso(centavos)
        partes.append(f"{ext_c} {'centavo' if centavos == 1 else 'centavos'}")
    if not partes:
        return "zero reais"
    return " e ".join(partes)


# ==============================================================================
# BANCO DE DADOS SQLITE (PROJETOS)
# ==============================================================================
def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS projetos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome_projeto TEXT,
            criado_em TEXT,
            atualizado_em TEXT,
            rito TEXT,
            proprietario TEXT,
            matricula TEXT,
            comarca TEXT,
            dados_json TEXT
        )
        """
    )
    conn.commit()
    conn.close()


def salvar_projeto(dados: dict) -> int:
    agora = datetime.now().strftime("%d/%m/%Y %H:%M")
    nome_proj = (dados.get("nome_projeto") or "").strip() or (
        f"{dados.get('imovel_denominacao') or 'Projeto Sem Nome'} — Mat. {dados.get('matricula_numero') or 'S/N'}"
    )
    dados["nome_projeto"] = nome_proj
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    proj_id = dados.get("projeto_id")
    if proj_id:
        cur.execute(
            """
            UPDATE projetos
            SET nome_projeto=?, atualizado_em=?, rito=?, proprietario=?, matricula=?, comarca=?, dados_json=?
            WHERE id=?
            """,
            (
                nome_proj, agora, dados.get("rito", ""), dados.get("proprietario_nome", ""),
                dados.get("matricula_numero", ""), dados.get("comarca_uf", ""),
                json.dumps(dados, ensure_ascii=False), proj_id,
            ),
        )
    else:
        cur.execute(
            """
            INSERT INTO projetos (nome_projeto, criado_em, atualizado_em, rito, proprietario, matricula, comarca, dados_json)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                nome_proj, agora, agora, dados.get("rito", ""), dados.get("proprietario_nome", ""),
                dados.get("matricula_numero", ""), dados.get("comarca_uf", ""),
                json.dumps(dados, ensure_ascii=False),
            ),
        )
        proj_id = cur.lastrowid
        dados["projeto_id"] = proj_id
    conn.commit()
    conn.close()
    return proj_id


def listar_projetos():
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute(
        "SELECT id, nome_projeto, rito, atualizado_em, dados_json FROM projetos ORDER BY id DESC LIMIT 100"
    ).fetchall()
    conn.close()
    return rows


def excluir_projeto(proj_id: int):
    conn = sqlite3.connect(DB_PATH)
    conn.execute("DELETE FROM projetos WHERE id=?", (proj_id,))
    conn.commit()
    conn.close()


def novo_projeto_vazio(nome: str = "", rito: str = None) -> dict:
    rito_final = rito or list(RITOS_MG.keys())[0]
    return {
        "projeto_id": None,
        "nome_projeto": nome,
        "rito": rito_final,
        "imovel_de_posse": "Posse" in rito_final or "Usucapião" in rito_final,
        "proprietario_tipo": "PF",
        "proprietario_nome": "",
        "proprietario_doc": "",
        "proprietario_rg": "",
        "proprietario_nac": "brasileiro(a)",
        "proprietario_estado_civil": "",
        "regime_bens": "",
        "proprietario_prof": "",
        "proprietario_filiacao": "",
        "proprietario_endereco": "",
        "proprietario_telefone": "",
        "proprietario_email": "",
        "conjuge_nome": "",
        "conjuge_cpf": "",
        "conjuge_rg": "",
        "conjuge_prof": "",
        "representante_legal_ou_inventariante": "",
        "imovel_denominacao": "",
        "imovel_municipio": "",
        "comarca_uf": "",
        "cartorio_cns": "",
        "matricula_numero": "",
        "livro_numero": "2",
        "incra_sncr": "",
        "cib_nirf": "",
        "car_numero": "",
        "sigef_codigo": "",
        "sigef_data": "",
        "art_numero": "",
        "area_matricula_ha": 0.0,
        "area_ccir_ha": 0.0,
        "area_memorial_ha": 0.0,
        "perimetro_m": 0.0,
        "fmp_municipio_ha": 3.0,
        "valor_mercado_num": "",
        "valor_mercado_extenso": "",
        "memorial_texto": "",
        "partes_memoriais": [
            {"nome": "Parte 1", "area_ha": 0.0, "perimetro_m": 0.0, "sigef": "", "valor_rs": "", "memorial_texto": ""}
        ],
        "confrontantes": [],
        "analise_base_feita": False,
    }


def normalizar_partes(proj: dict) -> list[dict]:
    """Garante compatibilidade entre projetos antigos (glebas/memorial_texto) e partes_memoriais."""
    partes = proj.get("partes_memoriais")
    if not partes:
        if proj.get("glebas"):
            partes = [
                {
                    "nome": g.get("nome") or f"Parte {i + 1}",
                    "area_ha": float(g.get("area_ha") or 0.0),
                    "perimetro_m": float(g.get("perimetro_m") or 0.0),
                    "sigef": g.get("sigef") or "",
                    "valor_rs": g.get("valor_rs") or "",
                    "memorial_texto": g.get("memorial_texto") or (proj.get("memorial_texto", "") if i == 0 else ""),
                }
                for i, g in enumerate(proj["glebas"])
            ]
        else:
            partes = [{
                "nome": "Parte 1",
                "area_ha": float(proj.get("area_memorial_ha") or 0.0),
                "perimetro_m": float(proj.get("perimetro_m") or 0.0),
                "sigef": proj.get("sigef_codigo") or "",
                "valor_rs": proj.get("valor_mercado_num") or "",
                "memorial_texto": proj.get("memorial_texto") or "",
            }]
        proj["partes_memoriais"] = partes
    return partes


# ==============================================================================
# LEITORES DE ARQUIVOS + MOTOR GEMINI 4 CARBON COM FALLBACK E RETRY 503
# ==============================================================================
def ler_rtf_bytes(file_bytes: bytes) -> str:
    raw = file_bytes.decode("latin-1", errors="ignore")
    raw = re.sub(r"\\'([0-9a-fA-F]{2})", lambda m: bytes.fromhex(m.group(1)).decode("cp1252", errors="ignore"), raw)
    raw = re.sub(r"\{\\(?:fonttbl|colortbl|stylesheet|info|pict|object)[^{}]*(?:\{[^{}]*\}[^{}]*)*\}", " ", raw)
    raw = re.sub(r"\\par[d]?|\\line", "\n", raw)
    raw = re.sub(r"\\[a-zA-Z]+-?\d* ?", " ", raw)
    raw = raw.replace("{", "").replace("}", "")
    raw = re.sub(r"[ \t]+", " ", raw)
    raw = re.sub(r"\n\s*\n+", "\n", raw)
    return raw.strip()


def ler_docx_bytes(file_bytes: bytes) -> str:
    d = docx.Document(io.BytesIO(file_bytes))
    partes = [p.text for p in d.paragraphs if p.text.strip()]
    for tbl in d.tables:
        for row in tbl.rows:
            partes.append(" | ".join(c.text.strip() for c in row.cells if c.text.strip()))
    return "\n".join(partes)


def ler_planilha_sigef(file_bytes: bytes, filename: str) -> str:
    engine = "odf" if filename.lower().endswith(".ods") else "openpyxl"
    try:
        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine=engine)
        resumo = [f"=== PLANILHA SIGEF/INCRA: {filename} ==="]
        for sheet in xls.sheet_names[:6]:
            df = xls.parse(sheet).dropna(how="all")
            resumo.append(f"--- Aba: {sheet} ({len(df)} linhas) ---")
            resumo.append(df.head(80).to_string(index=False))
        return "\n".join(resumo)
    except Exception as e:
        return f"[Aviso ao ler planilha {filename}: {e}]"


def preparar_partes_arquivos(arquivos_up: list, etiqueta: str) -> tuple[list, list]:
    binarios, textos = [], []
    for up in arquivos_up or []:
        raw = up.read()
        up.seek(0)
        fname = up.name.lower()
        if fname.endswith(".docx"):
            textos.append(f"=== [{etiqueta}] WORD: {up.name} ===\n{ler_docx_bytes(raw)}")
        elif fname.endswith(".rtf"):
            textos.append(f"=== [{etiqueta}] RTF (TopoCAD/AutoCAD): {up.name} ===\n{ler_rtf_bytes(raw)}")
        elif fname.endswith((".xlsx", ".ods")):
            textos.append(f"[{etiqueta}] " + ler_planilha_sigef(raw, up.name))
        elif fname.endswith(".txt"):
            textos.append(f"=== [{etiqueta}] TXT: {up.name} ===\n{raw.decode('utf-8', errors='ignore')}")
        elif fname.endswith(".pdf"):
            binarios.append(types.Part.from_bytes(data=raw, mime_type="application/pdf"))
            textos.append(f"[Arquivo PDF anexado em {etiqueta}: {up.name}]")
        elif fname.endswith((".png", ".jpg", ".jpeg", ".webp")):
            mime = "image/png" if fname.endswith(".png") else "image/jpeg"
            binarios.append(types.Part.from_bytes(data=raw, mime_type=mime))
            textos.append(f"[Imagem anexada em {etiqueta}: {up.name}]")
    return binarios, textos


MODELOS_GEMINI_OPCOES = {
    "Gemini 4 Carbon — Raciocínio Máximo (Mais Aprofundado)": "gemini-4-carbon",
    "Gemini 4 Neon — Inteligência de Fronteira + Velocidade": "gemini-4-neon",
    "Gemini 4 Argon — Raciocínio Eficiente": "gemini-4-argon",
    "Gemini 3.8 Flash — Rápido com Thinking": "gemini-3.8-flash",
}

MODELOS_GEMINI_FALLBACK = [
    "gemini-4-carbon",
    "gemini-4-neon",
    "gemini-4-argon",
    "gemini-4-hydrogen",
    "gemini-3.8-flash",
]


def chamar_gemini_json(api_key: str, partes_binarias: list, prompt: str, modelo_preferido: str = "gemini-4-carbon") -> dict:
    if not GENAI_AVAILABLE:
        raise RuntimeError("Biblioteca google-genai não instalada.")
    client = genai.Client(api_key=api_key)
    ordem = [modelo_preferido] + [m for m in MODELOS_GEMINI_FALLBACK if m != modelo_preferido]
    ultimo_erro = None
    resp = None
    modelo_usado = modelo_preferido
    config = types.GenerateContentConfig(
        response_mime_type="application/json",
        temperature=0.1,
    )
    erros_recuperaveis = (
        "404", "NOT_FOUND", "429", "RESOURCE_EXHAUSTED", "PERMISSION_DENIED",
        "503", "UNAVAILABLE", "500", "504", "high demand", "overloaded", "not available",
    )
    for modelo in ordem:
        for tentativa in range(2):
            try:
                resp = client.models.generate_content(
                    model=modelo,
                    contents=partes_binarias + [prompt],
                    config=config,
                )
                modelo_usado = modelo
                break
            except Exception as e:
                ultimo_erro = e
                err_s = str(e)
                if any(k in err_s for k in ("503", "UNAVAILABLE", "high demand", "overloaded")) and tentativa == 0:
                    time.sleep(2)
                    continue
                if any(k in err_s for k in erros_recuperaveis):
                    break
                raise
        if resp is not None:
            break
    if resp is None:
        raise RuntimeError(f"Nenhum modelo Gemini disponível respondeu. Último erro: {ultimo_erro}")
    st.session_state["ultimo_modelo_usado"] = modelo_usado
    txt = (resp.text or "").strip()
    txt = re.sub(r"^```json\s*", "", txt)
    txt = re.sub(r"\s*```$", "", txt)
    return json.loads(txt)


def extrair_base_e_confrontantes(api_key: str, rito: str, arq_prop, arq_imovel, arq_planta_mem, arq_comp, texto_obs: str, modelo_preferido: str = "gemini-4-carbon") -> dict:
    bins, txts = [], []
    for lista, tag in [
        (arq_prop, "PROPRIETÁRIO"),
        (arq_imovel, "IMÓVEL/CADASTROS"),
        (arq_planta_mem, "PLANTA E MEMORIAL AUTOCAD/TOPOCAD"),
        (arq_comp, "COMPLEMENTAR"),
    ]:
        b, t = preparar_partes_arquivos(lista, tag)
        bins.extend(b)
        txts.extend(t)
    if texto_obs.strip():
        txts.append(f"=== OBSERVAÇÕES / MEMORIAL COLADO ===\n{texto_obs.strip()}")

    prompt = f"""
Você é um analista documental cartorial sênior de Minas Gerais (Facillyte Topografia).
Rito selecionado: {rito}.
Leia TODOS os documentos do proprietário, do imóvel, da PLANTA e do(s) MEMORIAL(IS) DESCRITIVO(S) (AutoCAD / TopoCAD 2000 / SIGEF).

REGRAS CRÍTICAS DE EXTRAÇÃO:
1. MÚLTIPLAS PARTES OU GLEBAS: Se houver 2 ou mais partes, glebas, parcelas ou memoriais separados (ex.: "Parte 1" e "Parte 2", ou "Gleba 01", "Gleba 02" e "Remanescente"), retorne CADA UMA como um item distinto na lista "partes_memoriais", com seu nome exato, área (ha), perímetro (m), código SIGEF (se houver), valor (se houver) e o TEXTO INTEGRAL E COMPLETO do memorial descritivo daquela parte (com todos os vértices, coordenadas, azimutes, distâncias e confrontações, sem resumir!). Se houver apenas 1 polígono/parte, retorne 1 item em "partes_memoriais".
2. POSSE: Se o imóvel objeto não tiver matrícula (for posse/recibo) marque "imovel_de_posse": true. Se algum confrontante for posseiro/ocupante sem matrícula, marque "posseiro": true naquele confrontante.
3. CONFRONTANTES: Identifique TODOS os confrontantes citados em todas as plantas/memoriais (sem duplicatas), marcando "imovel_publico_ou_estrada": true se for estrada/rio/rua/área pública.

Retorne APENAS um JSON válido com estas chaves:
{{
  "imovel_de_posse": false,
  "proprietario_nome": "",
  "proprietario_tipo": "PF",
  "proprietario_doc": "",
  "proprietario_rg": "",
  "proprietario_nac": "brasileiro(a)",
  "proprietario_estado_civil": "",
  "regime_bens": "",
  "proprietario_prof": "",
  "proprietario_filiacao": "",
  "proprietario_endereco": "",
  "proprietario_telefone": "",
  "proprietario_email": "",
  "conjuge_nome": "",
  "conjuge_cpf": "",
  "conjuge_rg": "",
  "conjuge_prof": "",
  "representante_legal_ou_inventariante": "",
  "imovel_denominacao": "",
  "imovel_municipio": "",
  "comarca_uf": "",
  "cartorio_cns": "",
  "matricula_numero": "",
  "livro_numero": "2",
  "incra_sncr": "",
  "cib_nirf": "",
  "car_numero": "",
  "sigef_codigo": "",
  "sigef_data": "",
  "art_numero": "",
  "area_matricula_ha": 0.0,
  "area_ccir_ha": 0.0,
  "area_memorial_ha": 0.0,
  "perimetro_m": 0.0,
  "fmp_municipio_ha": 3.0,
  "valor_mercado_num": "",
  "valor_mercado_extenso": "",
  "memorial_texto": "",
  "partes_memoriais": [
    {{
      "nome": "Parte 1",
      "area_ha": 0.0,
      "perimetro_m": 0.0,
      "sigef": "",
      "valor_rs": "",
      "memorial_texto": ""
    }}
  ],
  "confrontantes": [
    {{
      "nome": "",
      "tipo": "PF",
      "doc": "",
      "rg": "",
      "qualificacao": "",
      "endereco": "",
      "matricula": "",
      "cns": "",
      "representante_ou_inventariante": "",
      "falecido": false,
      "posseiro": false,
      "imovel_publico_ou_estrada": false
    }}
  ]
}}
Não invente dados ausentes; deixe "" ou 0.0.
"""
    if txts:
        prompt += "\n\nTEXTOS E ARQUIVOS IDENTIFICADOS:\n" + "\n\n".join(txts)
    return chamar_gemini_json(api_key, bins, prompt, modelo_preferido=modelo_preferido)


def extrair_dados_confrontante(api_key: str, conf_atual: dict, arquivos_conf: list, texto_conf: str, modelo_preferido: str = "gemini-4-carbon") -> dict:
    bins, txts = preparar_partes_arquivos(arquivos_conf, f"CONFRONTANTE: {conf_atual.get('nome', '')}")
    if texto_conf.strip():
        txts.append(f"=== OBSERVAÇÕES DO CONFRONTANTE ===\n{texto_conf.strip()}")
    prompt = f"""
Você é um analista documental cartorial de MG.
Extraia os dados completos deste CONFRONTANTE a partir dos documentos anexados (RG, CPF, certidão, matrícula, recibo/declaração de posse, termo de inventariante/óbito ou contrato social).
Nome prévio: "{conf_atual.get('nome', '')}". Matrícula prévia: "{conf_atual.get('matricula', '')}".
Se o documento comprovar que o confrontante é posseiro/ocupante sem matrícula (recibo, contrato de compra e venda, cessão de posse), marque "posseiro": true.

Retorne APENAS um JSON válido com estas chaves:
{{
  "nome": "",
  "tipo": "PF",
  "doc": "",
  "rg": "",
  "qualificacao": "brasileiro(a), estado civil, profissão",
  "endereco": "",
  "matricula": "",
  "cns": "",
  "representante_ou_inventariante": "",
  "falecido": false,
  "posseiro": false,
  "imovel_publico_ou_estrada": false
}}
Não invente dados ausentes.
"""
    if txts:
        prompt += "\n\nTEXTOS EXTRAÍDOS:\n" + "\n\n".join(txts)
    return chamar_gemini_json(api_key, bins, prompt, modelo_preferido=modelo_preferido)


# ==============================================================================
# GUIA DINÂMICO DE DOCUMENTOS FÍSICOS PARA PROTOCOLO NO CRI (MG)
# ==============================================================================
def gerar_guia_protocolo_fisico_mg(d: dict) -> list[dict]:
    """Monta orientações legais e checklist de documentos físicos conforme o rito e as particularidades do caso (MG)."""
    rito = d.get("rito", "")
    info = RITOS_MG.get(rito, {})
    blocos = []

    # 1. Documentos físicos obrigatórios do Rito
    blocos.append({
        "titulo": f"📋 Documentos Físicos Obrigatórios — {rito}",
        "tipo": "base",
        "itens": list(info.get("docs_fisicos_base", [])),
    })

    # 2. Múltiplas partes / glebas
    partes = normalizar_partes(d)
    if len(partes) >= 2:
        blocos.append({
            "titulo": f"🗺️ Imóvel com {len(partes)} Partes / Glebas Distintas",
            "tipo": "destaque",
            "itens": [
                f"Protocolar os {len(partes)} memoriais descritivos completos em arquivos/vias separadas ({', '.join(p.get('nome', f'Parte {i+1}') for i, p in enumerate(partes))}), todos impressos em A4 e assinados com firma reconhecida.",
                "Conferir se a soma das áreas das partes/glebas bate com a área total informada no Requerimento e na Declaração de Valor Real e de Mercado (Lei Estadual nº 15.424/2004).",
                "Se houver certificação SIGEF por parcela, juntar a certificação impressa de cada parte.",
            ],
        })

    # 3. Imóvel de Posse (sem matrícula)
    if d.get("imovel_de_posse") or "posse" in (d.get("matricula_numero") or "").lower():
        blocos.append({
            "titulo": "📜 Atenção Especial — Imóvel de Posse (Sem Matrícula) em MG",
            "tipo": "alerta",
            "itens": [
                "Declaração de Posse, Recibo ou Contrato de Compra e Venda / Cessão de Posse: apresentar ORIGINAL ou CÓPIA AUTENTICADA com firma reconhecida das assinaturas (art. 861 do Prov. Conj. nº 93/2020/CGJ-MG).",
                "Cadastros rurais obrigatórios em MG: CCIR quitado (INCRA), CND do ITR (ou 5 últimos comprovantes de pagamento do ITR) e Recibo de Inscrição no CAR (SICAR/IEF-MG). *(Nota: o ITESP é órgão de São Paulo; em MG utilizam-se CCIR/INCRA, ITR e CAR.)*",
                "Para abertura de matrícula por Usucapião Extrajudicial (art. 216-A da Lei nº 6.015/73 e arts. 264 e 1.157 do Prov. Conj. nº 93/2020): juntar Ata Notarial lavrada em Tabelionato de Notas (mínimo 2 testemunhas se imóvel sem edificação), certidões negativas dos distribuidores (Estadual e Federal) e requerimento assinado por Advogado (OAB).",
            ],
        })

    # 4. Confrontantes Posseiros / Espólio / PJ
    confs = [c for c in d.get("confrontantes", []) if (c.get("nome") or "").strip() and not c.get("imovel_publico_ou_estrada")]
    posseiros = [c["nome"] for c in confs if c.get("posseiro") or "posse" in (c.get("matricula") or "").lower()]
    espolios = [c["nome"] for c in confs if c.get("falecido") or c.get("tipo") == "Espólio" or "espólio" in c["nome"].lower() or "espolio" in c["nome"].lower()]
    pjs = [c["nome"] for c in confs if c.get("tipo") == "PJ"]

    itens_esp = []
    if posseiros:
        itens_esp.append(
            f"Confrontante(s) Posseiro(s) ({', '.join(posseiros)}): juntar prova de posse/aquisição em ORIGINAL ou CÓPIA AUTENTICADA (contrato/recibo de compra e venda, termo/declaração de posse autenticada, CCIR, ITR, CAR ou conta de energia elétrica) + incluir declaração de buscas negativas de matrícula no requerimento (art. 213, §§ 10 e 17, Lei nº 6.015/73 e arts. 897–901 Prov. Conj. nº 93/2020)."
        )
    if espolios or d.get("proprietario_tipo") == "Espólio":
        itens_esp.append(
            f"Espólio ({', '.join(espolios) if espolios else d.get('proprietario_nome', 'Proprietário')}): se inventário aberto, juntar cópia autenticada do Termo de Inventariante (ou escritura pública) + Certidão de Óbito; se inventário não aberto, declaração de inventariante provisório + Certidão de Óbito + RG do inventariante."
        )
    if pjs or d.get("proprietario_tipo") == "PJ":
        itens_esp.append(
            f"Pessoa Jurídica ({', '.join(pjs) if pjs else d.get('proprietario_nome', 'Proprietária')}): juntar cópia autenticada do Contrato Social/Estatuto consolidado, última alteração/ata que comprove poderes do signatário e Cartão CNPJ atualizado."
        )
    if d.get("regime_bens") and "parcial" not in d.get("regime_bens", "").lower():
        itens_esp.append(
            f"Regime de bens ({d['regime_bens']}): apresentar Certidão de Registro do Pacto Antenupcial no Livro 3 (Registro Auxiliar) do CRI do domicílio dos cônjuges."
        )
    if itens_esp:
        blocos.append({
            "titulo": "⚖️ Documentos Complementares Específicos Detectados Neste Processo",
            "tipo": "alerta",
            "itens": itens_esp,
        })

    # 5. Regras de Autenticação Física vs. Eletrônica (MG)
    blocos.append({
        "titulo": "🏛️ Forma de Apresentação no Cartório (Físico ou Eletrônico — MG)",
        "tipo": "base",
        "itens": [
            "Protocolo Físico: peças técnicas e declarações originais com firma reconhecida em Tabelionato de Notas (dispensado se assinado presencialmente perante o Oficial/Escrevente — arts. 861 e 877 do Prov. Conj. nº 93/2020 c/c art. 3º, I, Lei nº 13.726/2018); cópias reprográficas devem ser autenticadas.",
            "Protocolo Eletrônico (CRI-MG / ONR): arquivos nato-digitais em PDF/A assinados com certificado ICP-Brasil (qualificada), gov.br (avançada), e-Notariado ou IdRC (art. 208, § 1º, I, e art. 329-A do Prov. nº 149/2023/CNJ).",
        ],
    })
    return blocos


# ==============================================================================
# MOTOR DE AUDITORIA ANTI-DEVOLUTIVA (MG)
# ==============================================================================
def auditar_processo(d: dict) -> list[dict]:
    alertas = []
    rito = d.get("rito", "")
    regra = RITOS_MG.get(rito, {})
    partes = normalizar_partes(d)

    if not d.get("proprietario_nome") or not d.get("proprietario_doc"):
        alertas.append({"nivel": "VERMELHO", "regra": "Qualificação do Proprietário", "detalhe": "Nome ou CPF/CNPJ do proprietário não informado (Prov. nº 61/2017/CNJ)."})
    else:
        alertas.append({"nivel": "VERDE", "regra": "Qualificação do Proprietário", "detalhe": f"{d['proprietario_nome']} ({d['proprietario_doc']})."})

    ec = (d.get("proprietario_estado_civil") or "").lower()
    if "casad" in ec or "união" in ec or "uniao" in ec:
        if not d.get("conjuge_nome") or not d.get("regime_bens"):
            alertas.append({"nivel": "VERMELHO", "regra": "Cônjuge e Regime de Bens", "detalhe": "Proprietário casado/união estável sem nome do cônjuge ou regime de bens."})
        else:
            alertas.append({"nivel": "VERDE", "regra": "Cônjuge e Regime de Bens", "detalhe": f"Cônjuge {d['conjuge_nome']} — {d['regime_bens']}."})
            if "parcial" not in (d.get("regime_bens") or "").lower():
                alertas.append({"nivel": "AMARELO", "regra": "Pacto Antenupcial", "detalhe": "Regime diverso da comunhão parcial: verificar registro do Pacto Antenupcial no Livro 3."})

    if d.get("proprietario_tipo") in ("PJ", "Espólio") and not d.get("representante_legal_ou_inventariante"):
        alertas.append({"nivel": "VERMELHO", "regra": "Representação Legal / Espólio", "detalhe": f"Proprietário é {d['proprietario_tipo']} sem representante legal/inventariante informado."})

    if d.get("imovel_de_posse") or "posse" in (d.get("matricula_numero") or "").lower():
        alertas.append({"nivel": "AMARELO", "regra": "Imóvel de Posse (Sem Matrícula)", "detalhe": "Anexar Declaração de Posse / Recibo autenticado com original, CCIR, ITR, CAR e conferir exigência de Ata Notarial (art. 264 Prov. Conj. 93/2020)."})

    if regra.get("exige_valor_mercado"):
        if not (d.get("valor_mercado_num") or "").strip():
            alertas.append({"nivel": "VERMELHO", "regra": "Valor Real e de Mercado (Lei Est. 15.424/2004)", "detalhe": "Obrigatório declarar valor real e de mercado (Tabela 4 de emolumentos em MG)."})
        else:
            alertas.append({"nivel": "VERDE", "regra": "Valor Real e de Mercado", "detalhe": f"R$ {d['valor_mercado_num']}."})

    # Conferência de múltiplas partes / glebas
    if len(partes) >= 2:
        sem_texto = [p.get("nome", f"Parte {i+1}") for i, p in enumerate(partes) if not (p.get("memorial_texto") or "").strip()]
        if sem_texto:
            alertas.append({"nivel": "AMARELO", "regra": "Memoriais por Parte/Gleba", "detalhe": f"Falta o texto completo do memorial em: {', '.join(sem_texto)}."})
        else:
            alertas.append({"nivel": "VERDE", "regra": "Memoriais Distintos por Parte/Gleba", "detalhe": f"{len(partes)} memoriais individuais prontos para geração em arquivos separados."})

    a_mat = float(d.get("area_matricula_ha") or 0)
    a_ccir = float(d.get("area_ccir_ha") or 0)
    soma_partes = sum(float(p.get("area_ha") or 0) for p in partes)
    a_mem = float(d.get("area_memorial_ha") or 0) or soma_partes

    if a_mat > 0 and a_mem > 0 and abs(a_mat - a_mem) > 0.0005:
        if rito == "Georreferenciamento (SIGEF/INCRA)":
            alertas.append({"nivel": "AMARELO", "regra": "Divergência Matrícula x Memorial", "detalhe": f"Matrícula ({a_mat:.4f} ha) difere do memorial ({a_mem:.4f} ha): acumular Georreferenciamento com Retificação (art. 891, § 9º, Prov. Conj. 93/2020)."})
        elif rito == "Inserção de Dados / Medidas Atualizadas":
            alertas.append({"nivel": "VERMELHO", "regra": "Alteração de Área na Inserção de Medidas", "detalhe": f"Memorial ({a_mem:.4f} ha) difere da matrícula ({a_mat:.4f} ha): mudar rito para Retificação Administrativa."})
        else:
            alertas.append({"nivel": "AMARELO", "regra": "Diferença de Área", "detalhe": f"Diferença de {abs(a_mat - a_mem):.4f} ha entre matrícula e memorial — conferir anuências."})
    elif a_mat > 0 and a_mem > 0:
        alertas.append({"nivel": "VERDE", "regra": "Compatibilidade de Áreas", "detalhe": f"Área do memorial ({a_mem:.4f} ha) compatível com a matrícula."})

    if a_ccir > 0 and a_mem > 0 and abs(a_ccir - a_mem) > 0.05:
        alertas.append({"nivel": "AMARELO", "regra": "CCIR x Memorial", "detalhe": f"Área do CCIR ({a_ccir:.4f} ha) difere do memorial ({a_mem:.4f} ha)."})

    if regra.get("exige_fmp"):
        fmp = float(d.get("fmp_municipio_ha") or 0)
        partes_validas = [p for p in partes if float(p.get("area_ha") or 0) > 0]
        if len(partes_validas) < 2:
            alertas.append({"nivel": "VERMELHO", "regra": "Glebas / Parcelas", "detalhe": "Cadastre todas as glebas e a área remanescente na seção de Partes/Glebas."})
        else:
            alvo = a_mem if a_mem > 0 else a_mat
            if alvo > 0 and abs(soma_partes - alvo) > 0.001:
                alertas.append({"nivel": "VERMELHO", "regra": "Soma de Áreas (Glebas + Remanescente)", "detalhe": f"Soma das partes ({soma_partes:.4f} ha) não fecha com a área total ({alvo:.4f} ha)."})
            else:
                alertas.append({"nivel": "VERDE", "regra": "Soma de Áreas", "detalhe": f"Soma das {len(partes_validas)} partes ({soma_partes:.4f} ha) fecha 100%."})
            abaixo = [p.get("nome", "Parcela") for p in partes_validas if fmp > 0 and float(p.get("area_ha") or 0) < fmp]
            if abaixo:
                alertas.append({"nivel": "VERMELHO", "regra": "Fração Mínima de Parcelamento (FMP)", "detalhe": f"Parcela(s) abaixo da FMP ({fmp:.4f} ha): {', '.join(abaixo)}."})
            elif fmp > 0:
                alertas.append({"nivel": "VERDE", "regra": "FMP", "detalhe": f"Todas as parcelas atendem à FMP ({fmp:.4f} ha)."})

    confs = [c for c in d.get("confrontantes", []) if (c.get("nome") or "").strip()]
    confs_privados = [c for c in confs if not c.get("imovel_publico_ou_estrada")]
    if regra.get("exige_confrontantes") and not confs_privados:
        alertas.append({"nivel": "VERMELHO", "regra": "Confrontantes", "detalhe": "Nenhum confrontante particular identificado/cadastrado."})

    mem_todos = " ".join([(p.get("memorial_texto") or "") for p in partes] + [d.get("memorial_texto") or ""]).lower()
    for c in confs_privados:
        nm = c.get("nome", "").strip()
        if not (c.get("doc") or "").strip():
            alertas.append({"nivel": "AMARELO", "regra": f"Documentos de {nm}", "detalhe": "CPF/CNPJ ainda não extraído — anexe na aba deste confrontante."})
        if c.get("posseiro") or "posse" in (c.get("matricula") or "").lower():
            alertas.append({"nivel": "AMARELO", "regra": f"Confrontante Posseiro ({nm})", "detalhe": "Juntar prova de posse autenticada (recibo/contrato/CCIR/ITR/CAR) + declaração de busca negativa de matrícula."})
        elif not (c.get("matricula") or "").strip():
            alertas.append({"nivel": "AMARELO", "regra": f"Matrícula de {nm}", "detalhe": "Matrícula/Posse do confrontante não informada."})
        if (c.get("falecido") or "espólio" in nm.lower() or "espolio" in nm.lower()) and not (c.get("representante_ou_inventariante") or "").strip():
            alertas.append({"nivel": "VERMELHO", "regra": f"Espólio ({nm})", "detalhe": "Confrontante falecido/espólio sem Inventariante informado."})
        if c.get("tipo") == "PJ" and not (c.get("representante_ou_inventariante") or "").strip():
            alertas.append({"nivel": "VERMELHO", "regra": f"Pessoa Jurídica ({nm})", "detalhe": "Confrontante PJ sem representante legal."})
        if mem_todos and nm.lower() not in mem_todos:
            alertas.append({"nivel": "AMARELO", "regra": f"Grafia no Memorial ({nm})", "detalhe": "Nome não aparece com a mesma grafia nos memoriais descritivos."})

    falt = [lbl for k, lbl in [("art_numero", "ART"), ("comarca_uf", "Comarca")] if not (d.get(k) or "").strip()]
    if not d.get("imovel_de_posse") and not (d.get("matricula_numero") or "").strip():
        falt.append("Matrícula")
    if falt:
        alertas.append({"nivel": "AMARELO", "regra": "Campos Essenciais", "detalhe": f"Preencher: {', '.join(falt)}."})
    else:
        alertas.append({"nivel": "VERDE", "regra": "Dados de Protocolo", "detalhe": "ART, Comarca e identificação registral preenchidos."})

    return alertas


# ==============================================================================
# GERADOR DE PEÇAS NO TIMBRADO OFICIAL DE 3 COLUNAS (.DOCX + .PDF)
# ==============================================================================
def cm_to_dxa(cm_val: float) -> int:
    return int(round(cm_val * 566.929))


def set_fixed_widths(table, widths_cm):
    table.autofit = False
    tblPr = table._tbl.tblPr
    total = sum(cm_to_dxa(w) for w in widths_cm)
    for tag in ("w:tblW", "w:tblLayout"):
        el = tblPr.find(qn(tag))
        if el is not None:
            tblPr.remove(el)
    tblPr.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="{total}" w:type="dxa"/>'))
    tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))
    old_grid = table._tbl.find(qn("w:tblGrid"))
    if old_grid is not None:
        table._tbl.remove(old_grid)
    grid = OxmlElement("w:tblGrid")
    for w in widths_cm:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(cm_to_dxa(w)))
        grid.append(gc)
    tblPr.addnext(grid)
    for row in table.rows:
        for i, w in enumerate(widths_cm):
            tcPr = row.cells[i]._tc.get_or_add_tcPr()
            old = tcPr.find(qn("w:tcW"))
            if old is not None:
                tcPr.remove(old)
            tcPr.append(parse_xml(f'<w:tcW {nsdecls("w")} w:w="{cm_to_dxa(w)}" w:type="dxa"/>'))


def set_cell_mar(cell, top=50, bottom=50, left=75, right=75):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m, v in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        n = OxmlElement(f"w:{m}")
        n.set(qn("w:w"), str(v))
        n.set(qn("w:type"), "dxa")
        tcMar.append(n)
    tcPr.append(tcMar)


def novo_doc_timbrado():
    doc = docx.Document()
    for sec in doc.sections:
        sec.page_width = Mm(210)
        sec.page_height = Mm(297)
        sec.top_margin = Cm(3.35)
        sec.bottom_margin = Cm(2.6)
        sec.left_margin = Cm(2.2)
        sec.right_margin = Cm(2.2)
        sec.header_distance = Cm(0.8)
        sec.footer_distance = Cm(0.85)

        hdr = sec.header
        hdr.is_linked_to_previous = False
        hdr.paragraphs[0].paragraph_format.line_spacing = Pt(1)
        htbl = hdr.add_table(rows=1, cols=3, width=Cm(16.6))
        htbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        htbl._tbl.tblPr.append(
            parse_xml(
                f'<w:tblBorders {nsdecls("w")}>'
                f'<w:top w:val="single" w:sz="8" w:space="0" w:color="{HEX_PRIMARY}"/>'
                f'<w:bottom w:val="single" w:sz="12" w:space="0" w:color="{HEX_PRIMARY}"/>'
                f'<w:left w:val="single" w:sz="4" w:space="0" w:color="{HEX_BORDER}"/>'
                f'<w:right w:val="single" w:sz="4" w:space="0" w:color="{HEX_BORDER}"/>'
                f'<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                f'<w:insideV w:val="single" w:sz="6" w:space="0" w:color="{HEX_BORDER}"/>'
                f'</w:tblBorders>'
            )
        )
        set_fixed_widths(htbl, [3.60, 9.70, 3.30])
        for idx in range(3):
            htbl.rows[0].cells[idx].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_mar(htbl.rows[0].cells[idx], 50, 50, 70, 70)

        p0 = htbl.rows[0].cells[0].paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if os.path.exists(LOGO_ENG):
            p0.add_run().add_picture(LOGO_ENG, width=Cm(2.85))
        else:
            r = p0.add_run("FACILLYTE\nENGENHARIA")
            r.font.bold, r.font.color.rgb = True, RGB_PRIMARY

        c1 = htbl.rows[0].cells[1]
        for idx, (txt, sz, bld, col) in enumerate([
            ("FACILLYTE ENGENHARIA LTDA", 12, True, RGB_PRIMARY),
            ("Topografia • Georreferenciamento • Regularização Fundiária", 9.5, True, RGB_ACCENT),
            ("Rua Ayrton Senna, 425 – Vila Nova • CEP 39.517-000 – Pai Pedro/MG", 9.5, False, RGB_SLATE),
            ("CNPJ: 40.310.224/0001-14 • Tel.: (38) 9 9881-0036 • facillyteengenharia@gmail.com", 9.5, False, RGB_MUTED),
        ]):
            pp = c1.paragraphs[0] if idx == 0 else c1.add_paragraph()
            pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pp.paragraph_format.space_before = Pt(0)
            pp.paragraph_format.space_after = Pt(1 if idx < 3 else 0)
            rr = pp.add_run(txt)
            rr.font.name, rr.font.size, rr.font.bold, rr.font.color.rgb = "Calibri", Pt(sz), bld, col

        p2 = htbl.rows[0].cells[2].paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if os.path.exists(LOGO_TOPO):
            p2.add_run().add_picture(LOGO_TOPO, width=Cm(2.15))
        else:
            r = p2.add_run("FACILLYTE\nTOPOGRAFIA")
            r.font.bold, r.font.color.rgb = True, RGB_ACCENT

        p_acc = hdr.add_paragraph()
        p_acc.paragraph_format.space_before, p_acc.paragraph_format.space_after, p_acc.paragraph_format.line_spacing = Pt(1), Pt(3), Pt(2)
        p_acc._p.get_or_add_pPr().append(
            parse_xml(f'<w:pBdr {nsdecls("w")}><w:top w:val="single" w:sz="8" w:space="0" w:color="{HEX_ACCENT}"/></w:pBdr>')
        )

        ftr = sec.footer
        ftr.is_linked_to_previous = False
        ftr.paragraphs[0].paragraph_format.line_spacing = Pt(1)
        ftbl = ftr.add_table(rows=1, cols=3, width=Cm(16.6))
        ftbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        ftbl._tbl.tblPr.append(
            parse_xml(
                f'<w:tblBorders {nsdecls("w")}>'
                f'<w:top w:val="single" w:sz="8" w:space="0" w:color="{HEX_PRIMARY}"/>'
                f'<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                f'<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                f'<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                f'<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                f'<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                f'</w:tblBorders>'
            )
        )
        set_fixed_widths(ftbl, [2.20, 12.20, 2.20])
        for idx in range(3):
            set_cell_mar(ftbl.rows[0].cells[idx], 50, 10, 15, 15)

        fc1 = ftbl.rows[0].cells[1]
        fp1 = fc1.paragraphs[0]
        fp1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = fp1.add_run("Rua Ayrton Senna, 425 – Vila Nova • CEP 39.517-000 – Pai Pedro/MG")
        r1.font.name, r1.font.size, r1.font.color.rgb = "Calibri", Pt(9.5), RGB_SLATE
        fp2 = fc1.add_paragraph()
        fp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r2 = fp2.add_run("CNPJ: 40.310.224/0001-14 • Tel.: (38) 9 9881-0036 • E-mail: facillyteengenharia@gmail.com")
        r2.font.name, r2.font.size, r2.font.color.rgb = "Calibri", Pt(9.5), RGB_MUTED

        fpr = ftbl.rows[0].cells[2].paragraphs[0]
        fpr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        rrp = fpr.add_run("Pág. ")
        rrp.font.name, rrp.font.size, rrp.font.color.rgb = "Calibri", Pt(9.5), RGB_MUTED
        fpr._p.append(parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="PAGE"/>'))
        rrs = fpr.add_run("/")
        rrs.font.name, rrs.font.size, rrs.font.color.rgb = "Calibri", Pt(9.5), RGB_MUTED
        fpr._p.append(parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="NUMPAGES"/>'))

    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)
    doc.styles["Normal"].font.color.rgb = RGB_SLATE
    return doc


def add_p(doc, text, bold=False, center=False, size=11, color=RGB_SLATE, space_after=6, justify=True):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else (WD_ALIGN_PARAGRAPH.JUSTIFY if justify else WD_ALIGN_PARAGRAPH.LEFT)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    r = p.add_run(text)
    r.font.name, r.font.size, r.font.bold, r.font.color.rgb = "Calibri", Pt(size), bold, color
    return p


def add_assinatura(doc, nome, legenda):
    p1 = add_p(doc, "____________________________________________________________", center=True, space_after=2)
    p1.paragraph_format.space_before = Pt(12)
    p1.paragraph_format.keep_with_next = True
    p2 = add_p(doc, nome, bold=True, center=True, space_after=1)
    p2.paragraph_format.keep_with_next = True
    add_p(doc, legenda, center=True, size=10, color=RGB_MUTED, space_after=6)


def qualificacao_proprietario(d: dict) -> str:
    partes = [d.get("proprietario_nome") or "[PROPRIETÁRIO]"]
    if d.get("proprietario_tipo") == "PJ":
        partes.append(f"pessoa jurídica inscrita no CNPJ nº {d.get('proprietario_doc', '')}")
        if d.get("representante_legal_ou_inventariante"):
            partes.append(f"neste ato representada por {d['representante_legal_ou_inventariante']}")
    else:
        for campo in [d.get("proprietario_nac"), d.get("proprietario_estado_civil"), d.get("proprietario_prof")]:
            if campo:
                partes.append(campo)
        if d.get("proprietario_rg"):
            partes.append(f"RG {d['proprietario_rg']}")
        if d.get("proprietario_doc"):
            partes.append(f"CPF {d['proprietario_doc']}")
        if d.get("conjuge_nome"):
            partes.append(f"casado(a) sob o regime de {d.get('regime_bens') or 'comunhão parcial de bens'} com {d['conjuge_nome']} (CPF {d.get('conjuge_cpf', '')})")
        if d.get("representante_legal_ou_inventariante"):
            partes.append(f"representado(a) por {d['representante_legal_ou_inventariante']}")
    if d.get("proprietario_endereco"):
        partes.append(f"residente/sediado(a) em {d['proprietario_endereco']}")
    return ", ".join(partes)


def gerar_pacote_zip(d: dict, alertas: list[dict], gerar_pdf: bool = True) -> bytes:
    data_hoje = datetime.now().strftime("%d/%m/%Y")
    comarca = d.get("comarca_uf") or "Porteirinha/MG"
    mun = d.get("imovel_municipio") or "Pai Pedro"
    rito = d.get("rito")
    fundamento = RITOS_MG[rito]["fundamento"]
    qual_prop = qualificacao_proprietario(d)
    rt_legenda = f"Engenheiro Civil • CREA-MG 248849/D • RNP 141908491-7/MG • INCRA: YIDB • ART nº {d.get('art_numero', '')}"
    partes = normalizar_partes(d)
    guia_fisico = gerar_guia_protocolo_fisico_mg(d)
    # Soma automática das partes se valor total estiver vazio + preenchimento automático do extenso
    val_total_num = parse_brl(d.get("valor_mercado_num"))
    if val_total_num is None and len(partes) >= 2:
        vals_p = [parse_brl(p.get("valor_rs")) for p in partes]
        if any(v is not None for v in vals_p):
            val_total_num = round(sum(v or 0.0 for v in vals_p), 2)
    val_num_fmt = formatar_brl(val_total_num) if val_total_num is not None else (d.get("valor_mercado_num") or "")
    val_ext_auto = valor_por_extenso_brl(val_num_fmt) or (d.get("valor_mercado_extenso") or "")

    docs_map = {}

    # 1. Requerimento Principal (citando automaticamente todas as partes/glebas e eventual posseiro)
    doc_req = novo_doc_timbrado()
    add_p(doc_req, f"ILUSTRÍSSIMO(A) SENHOR(A) OFICIAL(A) DO CARTÓRIO DE REGISTRO DE IMÓVEIS DA COMARCA DE {comarca.upper()}", bold=True, justify=False, space_after=8)
    add_p(doc_req, f"REQUERIMENTO DE {rito.upper()}", bold=True, center=True, size=13, color=RGB_PRIMARY, space_after=3)
    add_p(doc_req, fundamento, center=True, size=10, color=RGB_MUTED, space_after=8)
    mat_txt = "Imóvel de Posse (sem matrícula registrada)" if d.get("imovel_de_posse") else f"Matrícula nº {d.get('matricula_numero', '')}, Livro {d.get('livro_numero', '2')}, CNS {d.get('cartorio_cns', '')}"
    add_p(doc_req, f"{qual_prop}, juntamente com o Responsável Técnico MARCOS FRANCIEL FERNANDES DOS SANTOS, Engenheiro Civil, CREA-MG 248849/D, RNP 141908491-7/MG, INCRA YIDB, com escritório à Rua Ayrton Senna, 425, Vila Nova, Pai Pedro/MG, vêm requerer o processamento de {rito} no imóvel denominado {d.get('imovel_denominacao', '')}, Município de {mun}/MG, {mat_txt}, INCRA/SNCR {d.get('incra_sncr', '')}, CIB/NIRF {d.get('cib_nirf', '')}, CAR {d.get('car_numero', '')}, Certificação SIGEF {d.get('sigef_codigo', 'N/A')}.")

    if len(partes) >= 2:
        resumo_partes = "; ".join(
            f"{p.get('nome', f'Parte {i+1}')}: {float(p.get('area_ha') or 0):.4f} ha (Perímetro {float(p.get('perimetro_m') or 0):.3f} m"
            + (f", SIGEF {p['sigef']}" if p.get("sigef") else "") + ")"
            for i, p in enumerate(partes)
        )
        add_p(doc_req, f"Discriminação das {len(partes)} Partes / Glebas (memoriais individuais anexos): {resumo_partes}.", bold=True)

    add_p(doc_req, f"Área Registrada: {float(d.get('area_matricula_ha') or 0):.4f} ha | Área Total Apurada: {float(d.get('area_memorial_ha') or sum(float(p.get('area_ha') or 0) for p in partes)):.4f} ha | Valor Real e de Mercado (Lei Est. 15.424/2004, Tabela 4): R$ {val_num_fmt} ({val_ext_auto}).", bold=True)

    confs_posse = [c["nome"] for c in d.get("confrontantes", []) if c.get("posseiro") and (c.get("nome") or "").strip()]
    if confs_posse:
        add_p(doc_req, f"Declaração quanto a Confrontante(s) Posseiro(s) ({', '.join(confs_posse)}): Os requerentes e o Responsável Técnico declaram, nos termos do art. 213, §§ 10 e 17, da Lei nº 6.015/1973, que realizaram buscas e não localizaram matrícula nem titular de domínio do(s) imóvel(is) confrontante(s) ocupado(s) a título de posse, instruindo o pedido com a respectiva comprovação documental de posse autenticada.")

    add_p(doc_req, "Os requerentes e o Responsável Técnico declaram, sob responsabilidade civil e criminal (art. 213, § 14, da Lei nº 6.015/1973), que o levantamento foi realizado in loco (Datum SIRGAS 2000), respeitando integralmente as divisas existentes, faixas de domínio público e os direitos de todos os confrontantes, inexistindo invasão ou ação judicial conexa.")
    add_p(doc_req, f"Nestes termos, pede deferimento. {mun}/MG, {data_hoje}.", justify=False, space_after=8)
    add_assinatura(doc_req, d.get("proprietario_nome") or "[PROPRIETÁRIO]", f"Requerente — CPF/CNPJ: {d.get('proprietario_doc', '')}")
    if d.get("conjuge_nome"):
        add_assinatura(doc_req, d["conjuge_nome"], f"Cônjuge — CPF: {d.get('conjuge_cpf', '')}")
    add_assinatura(doc_req, "MARCOS FRANCIEL FERNANDES DOS SANTOS", rt_legenda)
    docs_map["01_Requerimento_CRI.docx"] = doc_req

    # 2. Anuências Individuais
    idx_c = 0
    for c in d.get("confrontantes", []):
        if not (c.get("nome") or "").strip() or c.get("imovel_publico_ou_estrada"):
            continue
        idx_c += 1
        doc_c = novo_doc_timbrado()
        add_p(doc_c, "DECLARAÇÃO DE ANUÊNCIA E RECONHECIMENTO DE LIMITES (CONFRONTANTE)", bold=True, center=True, size=13, color=RGB_PRIMARY, space_after=3)
        add_p(doc_c, "Art. 213, II e § 10, da Lei nº 6.015/1973 • Arts. 891 e 898 do Provimento Conjunto nº 93/2020 (CGJ/MG)", center=True, size=10, color=RGB_MUTED, space_after=8)
        rep = f", neste ato representado(a) por {c['representante_ou_inventariante']}" if c.get("representante_ou_inventariante") else ""
        cond_posse = "possuidor(a)/ocupante (comprovante de posse anexo)" if c.get("posseiro") else f"proprietário(a)/titular do imóvel Matrícula/Posse {c.get('matricula', '')} (CNS {c.get('cns', '')})"
        add_p(doc_c, f"Eu, {c['nome']}{rep}, {c.get('qualificacao', '')}, RG {c.get('rg', '')}, CPF/CNPJ {c.get('doc', '')}, residente/sediado(a) em {c.get('endereco', '')}, na qualidade de {cond_posse}, DECLARO, sob as penas da lei (art. 299 do Código Penal), que concordo integralmente com as medidas perimetrais, coordenadas, azimutes e divisas comuns constantes da planta e memorial(is) descritivo(s) do imóvel {d.get('imovel_denominacao', '')}, {mat_txt} (CRI {comarca}), de propriedade de {d.get('proprietario_nome', '')}, levantado pelo Eng. Civil Marcos Franciel Fernandes dos Santos (CREA-MG 248849/D, INCRA YIDB, ART nº {d.get('art_numero', '')}).")
        add_p(doc_c, "Declaro ciência de que, nos termos do art. 213, § 10, da Lei nº 6.015/1973 e art. 898 do Provimento Conjunto nº 93/2020/CGJ-MG, minha anuência supre a participação do cônjuge e de eventuais outros condôminos titulares do imóvel limítrofe.")
        add_p(doc_c, f"{mun}/MG, {data_hoje}.", justify=False, space_after=8)
        safe_name = re.sub(r"[^A-Za-z0-9]+", "_", c["nome"])[:30]
        add_assinatura(doc_c, c.get("representante_ou_inventariante") or c["nome"], f"Confrontante ({c['nome']}) — CPF/CNPJ: {c.get('doc', '')}")
        docs_map[f"02_Anuencia_{idx_c:02d}_{safe_name}.docx"] = doc_c

    # 3. Declaração Conjunta
    doc_dc = novo_doc_timbrado()
    add_p(doc_dc, "DECLARAÇÃO CONJUNTA DO(S) PROPRIETÁRIO(S) E DO RESPONSÁVEL TÉCNICO", bold=True, center=True, size=13, color=RGB_PRIMARY, space_after=3)
    add_p(doc_dc, "Art. 213, § 14, da Lei nº 6.015/1973 • Provimento Conjunto nº 93/2020 (CGJ/MG)", center=True, size=10, color=RGB_MUTED, space_after=8)
    add_p(doc_dc, f"Referência: Imóvel {d.get('imovel_denominacao', '')}, {mun}/MG, {mat_txt}, CRI de {comarca}.", bold=True)
    add_p(doc_dc, f"Responsável Técnico: Eu, MARCOS FRANCIEL FERNANDES DOS SANTOS, Engenheiro Civil, CREA-MG 248849/D, RNP 141908491-7/MG, INCRA YIDB (ART nº {d.get('art_numero', '')}), ATESTO sob as penas da lei que efetuei pessoalmente, in loco, o levantamento georreferenciado (SIRGAS 2000), sendo exatos os dados da planta e memorial(is) descritivo(s) e plenamente respeitadas as divisas particulares e faixas de domínio público.")
    add_p(doc_dc, f"Proprietário(s): {qual_prop}, ATESTAMOS sob as penas da lei serem verdadeiras todas as informações apresentadas, sem alteração de divisas nem invasão de áreas vizinhas, com plena ciência do art. 213, § 14, da Lei nº 6.015/1973 e do art. 299 do Código Penal.")
    add_p(doc_dc, f"{mun}/MG, {data_hoje}.", justify=False, space_after=8)
    add_assinatura(doc_dc, "MARCOS FRANCIEL FERNANDES DOS SANTOS", rt_legenda)
    add_assinatura(doc_dc, d.get("proprietario_nome") or "[PROPRIETÁRIO]", f"Proprietário(a) — CPF/CNPJ: {d.get('proprietario_doc', '')}")
    docs_map["03_Declaracao_Conjunta_Proprietario_RT.docx"] = doc_dc

    # 4. Declaração de Valor Real e de Mercado (com desdobramento por parte se 2+)
    doc_vm = novo_doc_timbrado()
    add_p(doc_vm, f"ILUSTRÍSSIMO(A) SENHOR(A) OFICIAL(A) DO CARTÓRIO DE REGISTRO DE IMÓVEIS DA COMARCA DE {comarca.upper()}", bold=True, justify=False, space_after=8)
    add_p(doc_vm, "DECLARAÇÃO DE VALOR REAL E DE MERCADO DO IMÓVEL", bold=True, center=True, size=13, color=RGB_PRIMARY, space_after=3)
    add_p(doc_vm, "Art. 10, § 3º, XVI, Lei Estadual nº 15.424/2004 (Tabela 4) • Arts. 134 a 138 do Prov. Conj. nº 93/2020 (CGJ/MG)", center=True, size=10, color=RGB_MUTED, space_after=8)
    add_p(doc_vm, f"{qual_prop}, proprietário(a) do imóvel {d.get('imovel_denominacao', '')}, {mat_txt}, CRI de {comarca}, requerente de {rito}, em atendimento aos arts. 134 a 138 do Provimento Conjunto nº 93/2020/CGJ-MG e ao art. 10, § 3º, XVI, da Lei Estadual nº 15.424/2004 (alterada pelas Leis nº 25.125/2024 e nº 25.367/2025), DECLARA, sob responsabilidade civil e penal, que o valor real e de mercado do imóvel é de R$ {val_num_fmt} ({val_ext_auto}).")
    if len(partes) >= 2:
        itens_det = []
        for i, p in enumerate(partes):
            vp = parse_brl(p.get("valor_rs"))
            vp_str = f"R$ {formatar_brl(vp)} ({valor_por_extenso_brl(vp)})" if vp is not None else "R$ [informar]"
            itens_det.append(f"{p.get('nome', f'Parte {i+1}')} ({float(p.get('area_ha') or 0):.4f} ha): {vp_str}")
        add_p(doc_vm, f"Discriminação por Parte / Gleba: {'; '.join(itens_det)}.", bold=True)
    add_p(doc_vm, "Declara ciência de que o valor destina-se ao recolhimento de emolumentos e TFJ, sujeito a arbitramento judicial se manifestamente inferior ao real (arts. 134 a 138 do Prov. Conj. 93/2020), bem como plena ciência do art. 299 do Código Penal Brasileiro.")
    add_p(doc_vm, f"{mun}/MG, {data_hoje}.", justify=False, space_after=8)
    add_assinatura(doc_vm, d.get("proprietario_nome") or "[PROPRIETÁRIO]", f"Proprietário(a) — CPF/CNPJ: {d.get('proprietario_doc', '')}")
    docs_map["04_Declaracao_Valor_Real_Mercado_MG.docx"] = doc_vm

    # 5. MEMORIAIS DESCRITIVOS DISTINTOS (1 ARQUIVO COMPLETO POR PARTE / GLEBA!)
    for idx_p, parte in enumerate(partes, start=1):
        p_nome = (parte.get("nome") or f"Parte {idx_p}").strip()
        p_area = float(parte.get("area_ha") or 0) or (float(d.get("area_memorial_ha") or 0) if len(partes) == 1 else 0.0)
        p_perim = float(parte.get("perimetro_m") or 0) or (float(d.get("perimetro_m") or 0) if len(partes) == 1 else 0.0)
        p_sigef = parte.get("sigef") or (d.get("sigef_codigo") if len(partes) == 1 else "")
        p_texto = (parte.get("memorial_texto") or "").strip() or (d.get("memorial_texto") or "[Inserir descrição completa do perímetro]")

        doc_mem = novo_doc_timbrado()
        titulo_mem = f"MEMORIAL DESCRITIVO — {p_nome.upper()}" if len(partes) >= 2 else "MEMORIAL DESCRITIVO"
        add_p(doc_mem, titulo_mem, bold=True, center=True, size=13.5, color=RGB_PRIMARY, space_after=4)
        add_p(doc_mem, f"Imóvel: {d.get('imovel_denominacao', '')}" + (f" ({p_nome})" if len(partes) >= 2 else "") + f" • {mat_txt} • Município: {mun}/MG • Comarca: {comarca}", bold=True, justify=False, space_after=4)
        add_p(doc_mem, f"Proprietário(a): {d.get('proprietario_nome', '')} ({d.get('proprietario_doc', '')}) • INCRA: {d.get('incra_sncr', '')}" + (f" • SIGEF: {p_sigef}" if p_sigef else "") + f" • Área: {p_area:.4f} ha • Perímetro: {p_perim:.3f} m • Datum: SIRGAS 2000", justify=False, space_after=8)
        add_p(doc_mem, p_texto)
        add_p(doc_mem, "Atesto, sob as penas da lei, que efetuei pessoalmente, in loco, o levantamento da área e que os valores corretos dos azimutes, distâncias, coordenadas e identificações dos confrontantes são os apresentados na planta e neste memorial descritivo, respeitadas as divisas particulares e faixas de domínio público.", space_after=6)
        add_p(doc_mem, f"{mun}/MG, {data_hoje}.", justify=False, space_after=8)
        add_assinatura(doc_mem, "MARCOS FRANCIEL FERNANDES DOS SANTOS", rt_legenda)
        add_assinatura(doc_mem, d.get("proprietario_nome") or "[PROPRIETÁRIO]", f"Proprietário(a) — CPF/CNPJ: {d.get('proprietario_doc', '')}")

        if len(partes) >= 2:
            safe_p = re.sub(r"[^A-Za-z0-9]+", "_", p_nome)[:25]
            fname_mem = f"05_Memorial_{idx_p:02d}_{safe_p}.docx"
        else:
            fname_mem = "05_Memorial_Descritivo_Timbrado.docx"
        docs_map[fname_mem] = doc_mem

    # 6. Checklist Anti-Devolutiva + Guia Completo de Documentos Físicos para Protocolo (MG)
    doc_chk = novo_doc_timbrado()
    add_p(doc_chk, "CHECKLIST ANTI-DEVOLUTIVA E GUIA DE PROTOCOLO FÍSICO (CRI — MG)", bold=True, center=True, size=13, color=RGB_PRIMARY, space_after=6)
    add_p(doc_chk, f"Projeto: {d.get('nome_projeto', '')} | Rito: {rito} | Comarca: {comarca} | Partes/Memoriais: {len(partes)}", bold=True, justify=False, space_after=6)

    add_p(doc_chk, "1. RESULTADO DA AUDITORIA AUTOMÁTICA", bold=True, color=RGB_PRIMARY, justify=False, space_after=4)
    for a in alertas:
        ico = {"VERDE": "[OK]", "AMARELO": "[ATENÇÃO]", "VERMELHO": "[BLOQUEIO]"}.get(a["nivel"], "")
        add_p(doc_chk, f"{ico} {a['regra']}: {a['detalhe']}", justify=False, space_after=3)

    add_p(doc_chk, "2. GUIA DE DOCUMENTOS FÍSICOS E ORIENTAÇÕES PARA PROTOCOLO NO CARTÓRIO (MG)", bold=True, color=RGB_PRIMARY, justify=False, space_after=4)
    for bloco in guia_fisico:
        add_p(doc_chk, bloco["titulo"], bold=True, color=RGB_ACCENT, justify=False, space_after=2)
        for item in bloco["itens"]:
            add_p(doc_chk, f"[  ] {item}", justify=True, space_after=3)

    docs_map["06_Checklist_e_Guia_Protocolo_CRI_MG.docx"] = doc_chk

    zip_buf = io.BytesIO()
    with tempfile.TemporaryDirectory() as tmpdir:
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for fname, d_obj in docs_map.items():
                docx_path = os.path.join(tmpdir, fname)
                d_obj.save(docx_path)
                zf.write(docx_path, arcname=f"DOCX_Editavel/{fname}")
                if gerar_pdf:
                    try:
                        res = subprocess.run(
                            ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", tmpdir, docx_path],
                            timeout=30,
                            capture_output=True,
                        )
                        pdf_name = fname.replace(".docx", ".pdf")
                        pdf_path = os.path.join(tmpdir, pdf_name)
                        if res.returncode == 0 and os.path.exists(pdf_path):
                            zf.write(pdf_path, arcname=f"PDF_Pronto_Impressao/{pdf_name}")
                    except Exception:
                        pass
    return zip_buf.getvalue()


# ==============================================================================
# INTERFACE PRINCIPAL (CORES FACILLYTE + FLUXO INTUITIVO POR PROJETO)
# ==============================================================================
init_db()
injetar_tema_facillyte()

if "projeto" not in st.session_state:
    st.session_state["projeto"] = None
if "etapa" not in st.session_state:
    st.session_state["etapa"] = 1

# ── BARRA LATERAL INSTITUCIONAL ──
with st.sidebar:
    st.markdown("### 📐 FACILLYTE DOCS")
    st.caption("Funcionário Virtual Cartorial • MG")

    if st.button("➕ Criar Novo Projeto", use_container_width=True, type="primary"):
        st.session_state["projeto"] = None
        st.session_state["etapa"] = 1
        st.rerun()

    st.divider()
    st.markdown("**📁 Projetos Salvos**")
    projetos = listar_projetos()
    if projetos:
        for pid, pnome, prito, patu, pjson in projetos:
            c_btn, c_del = st.columns([5, 1])
            ico = RITOS_MG.get(prito, {}).get("icone", "📄")
            if c_btn.button(f"{ico} {pnome}", key=f"open_{pid}", use_container_width=True, help=f"{prito} • {patu}"):
                st.session_state["projeto"] = json.loads(pjson)
                st.session_state["etapa"] = 2 if st.session_state["projeto"].get("analise_base_feita") else 1
                st.rerun()
            if c_del.button("🗑️", key=f"del_{pid}", help="Excluir projeto"):
                excluir_projeto(pid)
                if st.session_state["projeto"] and st.session_state["projeto"].get("projeto_id") == pid:
                    st.session_state["projeto"] = None
                st.rerun()
    else:
        st.caption("Seus projetos aparecerão aqui.")

    st.divider()
    secret_key = ""
    try:
        secret_key = st.secrets.get("GEMINI_API_KEY", "")
    except Exception:
        secret_key = ""
    api_key = st.text_input("🔑 Chave API Gemini", value=secret_key, type="password", help="Chave AIza... do Google AI Studio.")
    modelo_label = st.selectbox(
        "🧠 Motor de IA (Gemini)",
        list(MODELOS_GEMINI_OPCOES.keys()),
        index=0,
        help="Usa Gemini 4 Carbon (mais profundo) por padrão, com retentativa e fallback automático.",
    )
    modelo_selecionado = MODELOS_GEMINI_OPCOES[modelo_label]
    if st.session_state.get("ultimo_modelo_usado"):
        st.caption(f"Última leitura: `{st.session_state['ultimo_modelo_usado']}`")


# ── TELA INICIAL: CRIAR OU ABRIR PROJETO ──
if st.session_state["projeto"] is None:
    render_banner(
        "FACILLYTE ENGENHARIA & TOPOGRAFIA",
        "Funcionário Virtual de Peças Técnicas Cartoriais — Minas Gerais",
        "Provimento Conjunto nº 93/2020 • Lei nº 6.015/73 • Lei Est. nº 15.424/2004",
    )

    col_new, col_rec = st.columns([1.35, 1])
    with col_new:
        st.markdown(
            '<div class="fac-card fac-card-green"><h4>🚀 Iniciar Novo Projeto</h4>'
            '<p>Dê um nome ao processo e escolha o serviço. O sistema já monta automaticamente as caixas de documentos e as orientações legais de MG.</p></div>',
            unsafe_allow_html=True,
        )
        with st.form("form_novo_projeto"):
            nome_proj = st.text_input("Nome do Projeto", placeholder="Ex.: Fazenda Sonho Meu — Matrícula 6.724")
            rito_sel = st.selectbox(
                "Tipo de Serviço (Rito em MG)",
                list(RITOS_MG.keys()),
                format_func=lambda k: f"{RITOS_MG[k]['icone']} {k}",
            )
            st.caption(f"**Base legal:** {RITOS_MG[rito_sel]['fundamento']}")
            if st.form_submit_button("Criar Projeto e Começar ➡️", type="primary", use_container_width=True):
                proj = novo_projeto_vazio(nome=nome_proj.strip() or f"Projeto {datetime.now().strftime('%d/%m %H:%M')}", rito=rito_sel)
                salvar_projeto(proj)
                st.session_state["projeto"] = proj
                st.session_state["etapa"] = 1
                st.toast("Projeto criado com sucesso!", icon="✅")
                st.rerun()

    with col_rec:
        st.markdown(
            '<div class="fac-card"><h4>⚡ Como Funciona (4 Passos Simples)</h4>'
            '<p><b>1.</b> Envie docs do proprietário, imóvel, planta e memoriais (inclusive 2+ partes/glebas).<br>'
            '<b>2.</b> O sistema reconhece os confrontantes e abre uma aba para cada um.<br>'
            '<b>3.</b> Auditoria MG + Guia de Documentos Físicos (inclusive Posse/Espólio/PJ).<br>'
            '<b>4.</b> Baixe tudo no timbrado oficial com memoriais separados por parte (.docx + .pdf).</p></div>',
            unsafe_allow_html=True,
        )
        if projetos:
            st.markdown("**📂 Continuar Projeto Recente:**")
            for pid, pnome, prito, patu, pjson in projetos[:4]:
                if st.button(f"{RITOS_MG.get(prito, {}).get('icone', '📄')} {pnome} ({patu})", key=f"home_open_{pid}", use_container_width=True):
                    st.session_state["projeto"] = json.loads(pjson)
                    st.session_state["etapa"] = 2 if st.session_state["projeto"].get("analise_base_feita") else 1
                    st.rerun()
    st.stop()


# ── PROJETO ATIVO: CABEÇALHO + PROGRESSO VISUAL ──
proj = st.session_state["projeto"]
if proj.get("rito") not in RITOS_MG:
    proj["rito"] = list(RITOS_MG.keys())[0]
rito_info = RITOS_MG[proj["rito"]]
partes = normalizar_partes(proj)

render_banner(
    f"📂 {proj['nome_projeto']}",
    f"{rito_info['icone']} {proj['rito']} — Comarca: {proj.get('comarca_uf') or 'A definir'} • {len(partes)} parte(s)/memorial(is)",
    f"Passo {st.session_state['etapa']} de 4",
)
st.progress(st.session_state["etapa"] / 4.0)

etapas_labels = [
    "1️⃣ Documentos, Planta e Memoriais",
    f"2️⃣ Confrontantes ({len(proj.get('confrontantes', []))})",
    "3️⃣ Revisão, Partes & Guia CRI (MG)",
    "4️⃣ Baixar Peças Timbradas",
]
cols_nav = st.columns(4)
for idx_e, lbl in enumerate(etapas_labels, start=1):
    tipo_btn = "primary" if st.session_state["etapa"] == idx_e else "secondary"
    if cols_nav[idx_e - 1].button(lbl, key=f"nav_{idx_e}", use_container_width=True, type=tipo_btn):
        st.session_state["etapa"] = idx_e
        st.rerun()

st.write("")

# ==============================================================================
# PASSO 1: UPLOAD ORGANIZADO + LEITURA DE MÚLTIPLOS MEMORIAIS E CONFRONTANTES
# ==============================================================================
if st.session_state["etapa"] == 1:
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(
            f'<div class="fac-card"><h4>👤 1. Documentos do Proprietário</h4>'
            f'<p>{rito_info["docs_proprietario"]}</p></div>',
            unsafe_allow_html=True,
        )
        arq_prop = st.file_uploader(
            "Documentos do proprietário",
            type=["pdf", "png", "jpg", "jpeg", "webp", "docx"],
            accept_multiple_files=True,
            key="up_prop",
            label_visibility="collapsed",
        )
        if arq_prop:
            st.success(f"✅ {len(arq_prop)} arquivo(s) do proprietário anexado(s).")

        st.markdown(
            '<div class="fac-card fac-card-green"><h4>🗺️ 3. Planta e Memorial(is) Descritivo(s) (AutoCAD / TopoCAD 2000 / SIGEF)</h4>'
            '<p>Pode enviar 1 ou vários memoriais (ex.: Parte 1 e Parte 2, ou Glebas + Remanescente). O sistema separará cada memorial automaticamente e abrirá as abas dos confrontantes.</p></div>',
            unsafe_allow_html=True,
        )
        arq_planta_mem = st.file_uploader(
            "Planta e Memoriais",
            type=["pdf", "rtf", "docx", "txt", "png", "jpg", "jpeg", "xlsx", "ods"],
            accept_multiple_files=True,
            key="up_planta",
            label_visibility="collapsed",
        )
        if arq_planta_mem:
            st.success(f"✅ {len(arq_planta_mem)} arquivo(s) de planta/memorial anexado(s).")

    with c2:
        st.markdown(
            f'<div class="fac-card"><h4>🏡 2. Documentos do Imóvel e Cadastros</h4>'
            f'<p>{rito_info["docs_imovel"]}</p></div>',
            unsafe_allow_html=True,
        )
        arq_imovel = st.file_uploader(
            "Matrícula, CCIR, ITR, CAR, SIGEF e ART",
            type=["pdf", "png", "jpg", "jpeg", "webp", "docx", "xlsx", "ods"],
            accept_multiple_files=True,
            key="up_imovel",
            label_visibility="collapsed",
        )
        if arq_imovel:
            st.success(f"✅ {len(arq_imovel)} arquivo(s) do imóvel anexado(s).")

        st.markdown(
            '<div class="fac-card"><h4>📎 4. Complementares / Posse / Observações</h4>'
            '<p>Recibo ou declaração de posse, valor de mercado ou dados colados do WhatsApp do cliente.</p></div>',
            unsafe_allow_html=True,
        )
        arq_comp = st.file_uploader(
            "Documentos complementares",
            type=["pdf", "png", "jpg", "jpeg", "webp", "docx", "txt"],
            accept_multiple_files=True,
            key="up_comp",
            label_visibility="collapsed",
        )

    texto_obs = st.text_area(
        "📝 Texto livre (memorial do TopoCAD/AutoCAD, valor de mercado ou recados do WhatsApp)",
        value=proj.get("memorial_texto", ""),
        height=110,
    )

    col_go1, col_go2 = st.columns([3, 1])
    with col_go1:
        if st.button("🤖 Ler Tudo, Separar Partes/Memoriais e Identificar Confrontantes", type="primary", use_container_width=True):
            if not api_key:
                st.error("Informe a chave GEMINI_API_KEY na barra lateral (ou nos Secrets do Streamlit).")
            elif not (arq_prop or arq_imovel or arq_planta_mem or arq_comp or texto_obs.strip()):
                st.warning("Anexe pelo menos um documento ou cole o memorial antes de iniciar.")
            else:
                with st.spinner("Analisando documentos, separando partes/glebas e identificando confrontantes..."):
                    try:
                        ext = extrair_base_e_confrontantes(
                            api_key, proj["rito"], arq_prop, arq_imovel, arq_planta_mem, arq_comp, texto_obs,
                            modelo_preferido=modelo_selecionado,
                        )
                        ext["projeto_id"] = proj["projeto_id"]
                        ext["nome_projeto"] = proj["nome_projeto"]
                        ext["rito"] = proj["rito"]
                        ext["analise_base_feita"] = True
                        normalizar_partes(ext)
                        salvar_projeto(ext)
                        st.session_state["projeto"] = ext
                        st.session_state["etapa"] = 2
                        st.toast(
                            f"Pronto! {len(ext.get('partes_memoriais', []))} parte(s) e {len(ext.get('confrontantes', []))} confrontante(s) identificados.",
                            icon="🎉",
                        )
                        st.rerun()
                    except Exception as e:
                        err_msg = str(e)
                        if "API_KEY_INVALID" in err_msg or "API key not valid" in err_msg:
                            st.error("Sua chave GEMINI_API_KEY foi recusada pelo Gemini. Gere uma chave padrão ('AIza...') em aistudio.google.com/apikey.")
                        else:
                            st.error(f"Erro durante a leitura: {err_msg}")
    with col_go2:
        if st.button("Ir para Confrontantes ➡️", use_container_width=True):
            st.session_state["etapa"] = 2
            st.rerun()

# ==============================================================================
# PASSO 2: ABAS INDIVIDUAIS PARA CADA CONFRONTANTE RECONHECIDO
# ==============================================================================
elif st.session_state["etapa"] == 2:
    confs = proj.get("confrontantes", [])

    c_top1, c_top2 = st.columns([3, 1])
    with c_top1:
        if confs:
            st.markdown(
                f'<div class="fac-card fac-card-green"><h4>👥 {len(confs)} Confrontante(s) Identificado(s)</h4>'
                '<p>Clique em cada aba abaixo para anexar os documentos daquele confrontante (✅ completo • 🟡 pendente • 📜 posseiro • 🛣️ via pública).</p></div>',
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                '<div class="fac-card"><h4>👥 Nenhum confrontante listado ainda</h4>'
                '<p>Volte ao Passo 1 para enviar a planta e o memorial, ou clique em <b>+ Adicionar Confrontante</b> ao lado.</p></div>',
                unsafe_allow_html=True,
            )
    with c_top2:
        st.write("")
        if st.button("➕ Adicionar Confrontante", use_container_width=True, type="primary"):
            confs.append({
                "nome": f"Confrontante {len(confs) + 1}",
                "tipo": "PF", "doc": "", "rg": "", "qualificacao": "", "endereco": "",
                "matricula": "", "cns": "", "representante_ou_inventariante": "",
                "falecido": False, "posseiro": False, "imovel_publico_ou_estrada": False,
            })
            proj["confrontantes"] = confs
            salvar_projeto(proj)
            st.rerun()

    if confs:
        nomes_abas = []
        for i, c in enumerate(confs):
            if c.get("imovel_publico_ou_estrada"):
                status_ico = "🛣️"
            elif c.get("posseiro"):
                status_ico = "📜"
            elif (c.get("doc") or "").strip() and (c.get("matricula") or "").strip():
                status_ico = "✅"
            else:
                status_ico = "🟡"
            nomes_abas.append(f"{status_ico} {i + 1}. {(c.get('nome') or 'Sem nome')[:24]}")

        abas = st.tabs(nomes_abas)
        for i, aba in enumerate(abas):
            with aba:
                c = confs[i]
                col_up, col_form = st.columns([1, 1.35])
                with col_up:
                    st.markdown(
                        f'<div class="fac-card fac-card-green"><h4>📎 Documentos — {c.get("nome") or f"Confrontante {i + 1}"}</h4>'
                        f'<p>{rito_info["docs_confrontante"]}</p></div>',
                        unsafe_allow_html=True,
                    )
                    if c.get("posseiro"):
                        st.warning("📜 **Confrontante Posseiro (MG):** Para protocolo físico no CRI, junte prova de posse autenticada (recibo/contrato, termo de posse, CCIR, ITR, CAR ou conta de luz).")
                    arq_c = st.file_uploader(
                        f"Docs de {c.get('nome') or f'Confrontante {i + 1}'}",
                        type=["pdf", "png", "jpg", "jpeg", "webp", "docx", "txt"],
                        accept_multiple_files=True,
                        key=f"up_conf_{i}",
                        label_visibility="collapsed",
                    )
                    txt_c = st.text_area("Dados colados do WhatsApp deste confrontante (opcional)", height=85, key=f"txt_conf_{i}")
                    if st.button(f"🤖 Extrair dados de {c.get('nome') or f'Confrontante {i + 1}'}", key=f"btn_ext_conf_{i}", type="primary", use_container_width=True):
                        if not api_key:
                            st.error("Informe a chave GEMINI_API_KEY na barra lateral.")
                        elif not arq_c and not txt_c.strip():
                            st.warning("Anexe ao menos um documento ou cole os dados acima.")
                        else:
                            with st.spinner(f"Lendo documentos de {c.get('nome', 'confrontante')}..."):
                                try:
                                    novo_c = extrair_dados_confrontante(api_key, c, arq_c, txt_c, modelo_preferido=modelo_selecionado)
                                    for k, v in novo_c.items():
                                        if v not in ("", None):
                                            c[k] = v
                                    confs[i] = c
                                    proj["confrontantes"] = confs
                                    salvar_projeto(proj)
                                    st.toast("Dados do confrontante extraídos!", icon="✅")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Erro ao extrair: {e}")

                    if st.button(f"🗑️ Remover Confrontante {i + 1}", key=f"del_conf_{i}", use_container_width=True):
                        confs.pop(i)
                        proj["confrontantes"] = confs
                        salvar_projeto(proj)
                        st.rerun()

                with col_form:
                    st.markdown('<div class="fac-card"><h4>📋 Ficha do Confrontante</h4><p>Revise ou complete os dados abaixo.</p></div>', unsafe_allow_html=True)
                    c["nome"] = st.text_input("Nome / Razão Social / Espólio", value=c.get("nome", ""), key=f"cnome_{i}")
                    cc1, cc2, cc3, cc4 = st.columns(4)
                    tipos = ["PF", "PJ", "Espólio"]
                    c["tipo"] = cc1.selectbox("Tipo", tipos, index=tipos.index(c.get("tipo", "PF")) if c.get("tipo") in tipos else 0, key=f"ctipo_{i}")
                    c["posseiro"] = cc2.checkbox("Posseiro (s/ mat.)", value=bool(c.get("posseiro", False)), key=f"cpos_{i}")
                    c["falecido"] = cc3.checkbox("Falecido/Espólio", value=bool(c.get("falecido", False)), key=f"cfal_{i}")
                    c["imovel_publico_ou_estrada"] = cc4.checkbox("Via/Público", value=bool(c.get("imovel_publico_ou_estrada", False)), key=f"cpub_{i}")
                    cc5, cc6 = st.columns(2)
                    c["doc"] = cc5.text_input("CPF / CNPJ", value=c.get("doc", ""), key=f"cdoc_{i}")
                    c["rg"] = cc6.text_input("RG / Órgão", value=c.get("rg", ""), key=f"crg_{i}")
                    c["qualificacao"] = st.text_input("Qualificação (nacionalidade, estado civil, profissão)", value=c.get("qualificacao", ""), key=f"cqual_{i}")
                    c["endereco"] = st.text_input("Endereço Completo", value=c.get("endereco", ""), key=f"cend_{i}")
                    cc7, cc8 = st.columns(2)
                    c["matricula"] = cc7.text_input("Matrícula / Transcrição / Posse", value=c.get("matricula", ""), key=f"cmat_{i}")
                    c["cns"] = cc8.text_input("CNS do Cartório", value=c.get("cns", ""), key=f"ccns_{i}")
                    c["representante_ou_inventariante"] = st.text_input("Representante Legal / Inventariante (se PJ ou Espólio)", value=c.get("representante_ou_inventariante", ""), key=f"crep_{i}")
                    confs[i] = c

    proj["confrontantes"] = confs
    st.write("")
    c_prev, c_save, c_next = st.columns([1, 1, 2])
    if c_prev.button("⬅️ Voltar ao Passo 1", use_container_width=True):
        salvar_projeto(proj)
        st.session_state["etapa"] = 1
        st.rerun()
    if c_save.button("💾 Salvar Alterações", use_container_width=True):
        salvar_projeto(proj)
        st.toast("Alterações salvas!", icon="💾")
    if c_next.button("Avançar para Revisão, Partes e Guia CRI ➡️", type="primary", use_container_width=True):
        salvar_projeto(proj)
        st.session_state["etapa"] = 3
        st.rerun()

# ==============================================================================
# PASSO 3: REVISÃO + MÚLTIPLAS PARTES/MEMORIAIS + AUDITORIA + GUIA FÍSICO MG
# ==============================================================================
elif st.session_state["etapa"] == 3:
    alertas = auditar_processo(proj)
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("🔴 Bloqueios", sum(1 for a in alertas if a["nivel"] == "VERMELHO"))
    m2.metric("🟡 Atenção", sum(1 for a in alertas if a["nivel"] == "AMARELO"))
    m3.metric("🟢 Validados", sum(1 for a in alertas if a["nivel"] == "VERDE"))
    m4.metric("📄 Memoriais/Partes", len(partes))

    with st.expander("🔎 Painel de Auditoria Anti-Devolutiva (Provimento Conjunto nº 93/2020 — MG)", expanded=True):
        for a in alertas:
            if a["nivel"] == "VERMELHO":
                st.error(f"**{a['regra']}:** {a['detalhe']}")
            elif a["nivel"] == "AMARELO":
                st.warning(f"**{a['regra']}:** {a['detalhe']}")
            else:
                st.success(f"**{a['regra']}:** {a['detalhe']}")

    # Guia Legal Dinâmico de Documentos Físicos para Protocolo no CRI (MG)
    guia_fisico = gerar_guia_protocolo_fisico_mg(proj)
    with st.expander("🏛️ Orientação Legal: Documentos Físicos Obrigatórios para Protocolo no Cartório (MG)", expanded=True):
        for bloco in guia_fisico:
            cls = "fac-card fac-card-amber" if bloco["tipo"] == "alerta" else ("fac-card fac-card-green" if bloco["tipo"] == "destaque" else "fac-card")
            itens_html = "<br>".join(f"• {it}" for it in bloco["itens"])
            st.markdown(f'<div class="{cls}"><h4>{bloco["titulo"]}</h4><p>{itens_html}</p></div>', unsafe_allow_html=True)

    # Seção de Partes / Glebas e Memoriais Distintos
    c_ph1, c_ph2 = st.columns([3, 1])
    with c_ph1:
        st.markdown(
            f'<div class="fac-card fac-card-green"><h4>🗺️ Partes / Glebas e Memoriais Descritivos ({len(partes)})</h4>'
            '<p>Cada aba abaixo gera um arquivo de Memorial Descritivo completo e separado no pacote final (.docx e .pdf).</p></div>',
            unsafe_allow_html=True,
        )
    with c_ph2:
        st.write("")
        if st.button("➕ Adicionar Parte / Gleba", use_container_width=True, type="primary"):
            partes.append({
                "nome": f"Parte {len(partes) + 1}",
                "area_ha": 0.0, "perimetro_m": 0.0, "sigef": "", "valor_rs": "", "memorial_texto": "",
            })
            proj["partes_memoriais"] = partes
            salvar_projeto(proj)
            st.rerun()

    abas_p = st.tabs([f"📄 {p.get('nome') or f'Parte {i+1}'}" for i, p in enumerate(partes)])
    for i, ab in enumerate(abas_p):
        with ab:
            p = partes[i]
            cp1, cp2, cp3, cp4, cp5 = st.columns([1.4, 1, 1, 1.3, 1])
            p["nome"] = cp1.text_input("Nome da Parte / Gleba", value=p.get("nome", f"Parte {i+1}"), key=f"pnome_{i}")
            p["area_ha"] = cp2.number_input("Área (ha)", value=float(p.get("area_ha") or 0.0), format="%.4f", key=f"parea_{i}")
            p["perimetro_m"] = cp3.number_input("Perímetro (m)", value=float(p.get("perimetro_m") or 0.0), format="%.3f", key=f"pper_{i}")
            p["sigef"] = cp4.text_input("Código SIGEF da Parte", value=p.get("sigef", ""), key=f"psig_{i}")
            p["valor_rs"] = cp5.text_input("Valor da Parte (R$)", value=p.get("valor_rs", ""), key=f"pval_{i}")
            p["memorial_texto"] = st.text_area(
                f"Memorial Descritivo Completo — {p['nome']}",
                value=p.get("memorial_texto", ""),
                height=160,
                key=f"pmem_{i}",
            )
            if len(partes) > 1 and st.button(f"🗑️ Remover {p['nome']}", key=f"pdel_{i}"):
                partes.pop(i)
                proj["partes_memoriais"] = partes
                salvar_projeto(proj)
                st.rerun()
            partes[i] = p
    proj["partes_memoriais"] = partes

    st.markdown('<div class="fac-card"><h4>📝 Conferência Rápida — Proprietário e Imóvel</h4><p>Revise os dados abaixo antes de gerar o pacote final.</p></div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        proj["proprietario_tipo"] = st.selectbox("Tipo de Proprietário", ["PF", "PJ", "Espólio"], index=["PF", "PJ", "Espólio"].index(proj.get("proprietario_tipo", "PF")))
        proj["proprietario_nome"] = st.text_input("Nome / Razão Social", value=proj.get("proprietario_nome", ""))
        proj["proprietario_doc"] = st.text_input("CPF / CNPJ", value=proj.get("proprietario_doc", ""))
        proj["proprietario_rg"] = st.text_input("RG / Órgão Expedidor", value=proj.get("proprietario_rg", ""))
        proj["proprietario_estado_civil"] = st.text_input("Estado Civil", value=proj.get("proprietario_estado_civil", ""))
        proj["regime_bens"] = st.text_input("Regime de Bens", value=proj.get("regime_bens", ""))
        proj["proprietario_prof"] = st.text_input("Profissão", value=proj.get("proprietario_prof", ""))
        proj["proprietario_endereco"] = st.text_input("Endereço Completo", value=proj.get("proprietario_endereco", ""))
        proj["conjuge_nome"] = st.text_input("Nome do Cônjuge (se casado)", value=proj.get("conjuge_nome", ""))
        proj["conjuge_cpf"] = st.text_input("CPF do Cônjuge", value=proj.get("conjuge_cpf", ""))
        proj["representante_legal_ou_inventariante"] = st.text_input("Representante Legal / Inventariante", value=proj.get("representante_legal_ou_inventariante", ""))

    with col2:
        proj["imovel_de_posse"] = st.checkbox("📜 Imóvel de Posse (sem matrícula registrada — ativa orientações de posse/autenticação)", value=bool(proj.get("imovel_de_posse", False)))
        proj["imovel_denominacao"] = st.text_input("Denominação do Imóvel", value=proj.get("imovel_denominacao", ""))
        proj["imovel_municipio"] = st.text_input("Município (MG)", value=proj.get("imovel_municipio", ""))
        proj["comarca_uf"] = st.text_input("Comarca do CRI (ex.: Porteirinha/MG)", value=proj.get("comarca_uf", ""))
        proj["cartorio_cns"] = st.text_input("CNS do Cartório", value=proj.get("cartorio_cns", ""))
        proj["matricula_numero"] = st.text_input("Nº da Matrícula / Transcrição", value=proj.get("matricula_numero", ""))
        proj["incra_sncr"] = st.text_input("CCIR / SNCR", value=proj.get("incra_sncr", ""))
        proj["cib_nirf"] = st.text_input("CIB / NIRF", value=proj.get("cib_nirf", ""))
        proj["car_numero"] = st.text_input("Recibo CAR", value=proj.get("car_numero", ""))
        proj["sigef_codigo"] = st.text_input("Código(s) SIGEF", value=proj.get("sigef_codigo", ""))
        proj["art_numero"] = st.text_input("Nº da ART (CREA-MG)", value=proj.get("art_numero", ""))
        ca1, ca2, ca3 = st.columns(3)
        proj["area_matricula_ha"] = ca1.number_input("Área Matrícula (ha)", value=float(proj.get("area_matricula_ha") or 0.0), format="%.4f")
        proj["area_ccir_ha"] = ca2.number_input("Área CCIR (ha)", value=float(proj.get("area_ccir_ha") or 0.0), format="%.4f")
        proj["area_memorial_ha"] = ca3.number_input("Área Total Apurada (ha)", value=float(proj.get("area_memorial_ha") or sum(float(p.get("area_ha") or 0) for p in partes)), format="%.4f")
        cp1, cp2 = st.columns(2)
        proj["perimetro_m"] = cp1.number_input("Perímetro Total (m)", value=float(proj.get("perimetro_m") or sum(float(p.get("perimetro_m") or 0) for p in partes)), format="%.3f")
        proj["fmp_municipio_ha"] = cp2.number_input("FMP Município (ha)", value=float(proj.get("fmp_municipio_ha") or 3.0), format="%.4f")
        cv1, cv2 = st.columns([1, 1.4])
        val_inicial = proj.get("valor_mercado_num", "")
        if not val_inicial and len(partes) >= 2:
            vals_p = [parse_brl(p.get("valor_rs")) for p in partes]
            if any(v is not None for v in vals_p):
                val_inicial = formatar_brl(sum(v or 0.0 for v in vals_p))
        proj["valor_mercado_num"] = cv1.text_input(
            "Valor Real e de Mercado (R$)",
            value=val_inicial,
            placeholder="Ex.: 150.165,00",
            help="Digite apenas o número — o valor por extenso é preenchido automaticamente.",
        )
        extenso_auto = valor_por_extenso_brl(proj["valor_mercado_num"])
        proj["valor_mercado_extenso"] = extenso_auto or proj.get("valor_mercado_extenso", "")
        cv2.text_input(
            "Valor por Extenso (automático)",
            value=proj["valor_mercado_extenso"],
            disabled=True,
        )

    c_b1, c_b2, c_b3 = st.columns([1, 1, 2])
    if c_b1.button("⬅️ Voltar aos Confrontantes", use_container_width=True):
        salvar_projeto(proj)
        st.session_state["etapa"] = 2
        st.rerun()
    if c_b2.button("🔄 Reauditar e Salvar", use_container_width=True):
        salvar_projeto(proj)
        st.toast("Auditoria e guia atualizados!", icon="🔄")
        st.rerun()
    if c_b3.button("Avançar para Gerar Peças Timbradas ➡️", type="primary", use_container_width=True):
        salvar_projeto(proj)
        st.session_state["etapa"] = 4
        st.rerun()

# ==============================================================================
# PASSO 4: GERAR PACOTE COMPLETO NO TIMBRADO OFICIAL (.ZIP)
# ==============================================================================
elif st.session_state["etapa"] == 4:
    alertas_finais = auditar_processo(proj)
    if any(a["nivel"] == "VERMELHO" for a in alertas_finais):
        st.warning("Existem bloqueios vermelhos na auditoria (Passo 3). Você pode gerar agora como rascunho ou voltar para corrigir antes do protocolo.")

    confs_priv = [c for c in proj.get("confrontantes", []) if (c.get("nome") or "").strip() and not c.get("imovel_publico_ou_estrada")]
    lista_mems = ", ".join(p.get("nome", f"Parte {i+1}") for i, p in enumerate(partes))
    st.markdown(
        f'<div class="fac-card fac-card-green"><h4>📦 Dossiê Pronto para Geração no Timbrado Oficial Facillyte</h4>'
        f'<p>• 1 Requerimento de <b>{proj["rito"]}</b> para o CRI de <b>{proj.get("comarca_uf") or "[Comarca]"}</b><br>'
        f'• <b>{len(confs_priv)}</b> Declaração(ões) individual(is) de anuência de confrontante<br>'
        f'• 1 Declaração Conjunta (Proprietário + Eng. Civil Marcos Franciel, CREA-MG 248849/D)<br>'
        f'• 1 Declaração de Valor Real e de Mercado (Lei Estadual nº 15.424/2004)<br>'
        f'• <b>{len(partes)} Memorial(is) Descritivo(s) distinto(s)</b> ({lista_mems})<br>'
        f'• 1 Checklist Anti-Devolutiva + <b>Guia Completo de Documentos Físicos para Protocolo (MG)</b></p></div>',
        unsafe_allow_html=True,
    )

    gerar_pdf = st.checkbox("Incluir também os arquivos em PDF prontos para impressão além do Word (.docx)", value=True)
    if st.button("🚀 Gerar Pacote Completo para Protocolo (.zip)", type="primary", use_container_width=True):
        salvar_projeto(proj)
        with st.spinner("Montando todas as peças e memoriais individuais no timbrado oficial..."):
            zip_bytes = gerar_pacote_zip(proj, alertas_finais, gerar_pdf=gerar_pdf)
            nome_arq = re.sub(r"[^A-Za-z0-9]+", "_", f"Dossie_{proj.get('nome_projeto') or 'Facillyte'}") + ".zip"
            st.toast("Pacote gerado com sucesso!", icon="🎉")
            st.download_button(
                label=f"⬇️ Baixar {nome_arq}",
                data=zip_bytes,
                file_name=nome_arq,
                mime="application/zip",
                use_container_width=True,
            )

    if st.button("⬅️ Voltar para Revisão e Guia CRI"):
        st.session_state["etapa"] = 3
        st.rerun()