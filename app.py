import io
import json
import os
import re
import sqlite3
import subprocess
import tempfile
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
# CONFIGURAÇÃO INICIAL E IDENTIDADE FACILLYTE
# ==============================================================================
st.set_page_config(
    page_title="Facillyte Docs — Funcionário Virtual Cartorial (MG)",
    page_icon="📐",
    layout="wide",
)

DB_PATH = "facillyte_processos.db"
LOGO_ENG = os.path.join("assets", "logo_facillyte_engenharia.png")
LOGO_TOPO = os.path.join("assets", "logo_facillyte_topografia.png")

HEX_PRIMARY = "0B2457"
HEX_ACCENT = "0A8A1C"
HEX_SLATE = "1E293B"
HEX_MUTED = "475569"
HEX_TINT = "F1F5F9"
HEX_BORDER = "CBD5E1"

RGB_PRIMARY = RGBColor(0x0B, 0x24, 0x57)
RGB_ACCENT = RGBColor(0x0A, 0x8A, 0x1C)
RGB_SLATE = RGBColor(0x1E, 0x29, 0x3B)
RGB_MUTED = RGBColor(0x47, 0x55, 0x69)
RGB_WHITE = RGBColor(0xFF, 0xFF, 0xFF)

RITOS_MG = {
    "Retificação Administrativa de Área": {
        "fundamento": "Art. 213, II, da Lei nº 6.015/1973 e arts. 889 a 900 (art. 891) do Provimento Conjunto nº 93/2020 (CGJ/MG)",
        "exige_confrontantes": True,
        "exige_fmp": False,
        "exige_valor_mercado": True,
    },
    "Desmembramento Rural": {
        "fundamento": "Arts. 176 e 246 da Lei nº 6.015/1973, art. 65 da Lei nº 4.504/1964 e art. 991 do Provimento Conjunto nº 93/2020 (CGJ/MG)",
        "exige_confrontantes": False,
        "exige_fmp": True,
        "exige_valor_mercado": True,
    },
    "Estremação (Condomínio Pro Diviso)": {
        "fundamento": "Arts. 1.149 a 1.155 do Provimento Conjunto nº 93/2020 (CGJ/MG) e art. 65 da Lei nº 4.504/1964",
        "exige_confrontantes": True,
        "exige_fmp": True,
        "exige_valor_mercado": True,
    },
    "Inserção de Dados / Medidas Atualizadas": {
        "fundamento": "Art. 213, I e II, da Lei nº 6.015/1973 e art. 891 do Provimento Conjunto nº 93/2020 (CGJ/MG)",
        "exige_confrontantes": True,
        "exige_fmp": False,
        "exige_valor_mercado": True,
    },
    "Georreferenciamento (SIGEF/INCRA)": {
        "fundamento": "Art. 176, §§ 3º a 5º, Lei nº 6.015/1973, Lei nº 10.267/2001, Decretos nº 4.449/2002 e 5.570/2005 e arts. 891 (§ 9º) e 1.027+ do Prov. Conj. nº 93/2020",
        "exige_confrontantes": True,
        "exige_fmp": False,
        "exige_valor_mercado": True,
    },
    "Averbação Diversa (CAR / CCIR / Construção / Subjetiva)": {
        "fundamento": "Arts. 167, II, e 246 da Lei nº 6.015/1973, art. 717 do Provimento Conjunto nº 93/2020 (CGJ/MG) e Provimento nº 61/2017/CNJ",
        "exige_confrontantes": False,
        "exige_fmp": False,
        "exige_valor_mercado": False,
    },
}


# ==============================================================================
# BANCO DE DADOS SQLITE (HISTÓRICO DE CLIENTES E PROCESSOS)
# ==============================================================================
def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS processos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            criado_em TEXT,
            rito TEXT,
            proprietario TEXT,
            imovel TEXT,
            matricula TEXT,
            comarca TEXT,
            dados_json TEXT
        )
        """
    )
    conn.commit()
    conn.close()


def salvar_processo(dados: dict):
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        """
        INSERT INTO processos (criado_em, rito, proprietario, imovel, matricula, comarca, dados_json)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (
            datetime.now().strftime("%d/%m/%Y %H:%M"),
            dados.get("rito", ""),
            dados.get("proprietario_nome", ""),
            dados.get("imovel_denominacao", ""),
            dados.get("matricula_numero", ""),
            dados.get("comarca_uf", ""),
            json.dumps(dados, ensure_ascii=False),
        ),
    )
    conn.commit()
    conn.close()


def listar_processos():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute(
        "SELECT id, criado_em, rito, proprietario, imovel, matricula, comarca, dados_json "
        "FROM processos ORDER BY id DESC LIMIT 50"
    )
    rows = cur.fetchall()
    conn.close()
    return rows


# ==============================================================================
# LEITURA DE ARQUIVOS HETEROGÊNEOS (WORD, EXCEL/ODS SIGEF, PDF/IMAGENS)
# ==============================================================================
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
            resumo.append(df.head(60).to_string(index=False))
        return "\n".join(resumo)
    except Exception as e:
        return f"[Aviso ao ler planilha {filename}: {e}]"


def extrair_com_gemini(api_key: str, arquivos_up: list, texto_livre: str) -> dict:
    if not GENAI_AVAILABLE:
        raise RuntimeError("Biblioteca google-genai não instalada.")
    client = genai.Client(api_key=api_key)

    partes_conteudo = []
    textos_extraidos = []

    if texto_livre.strip():
        textos_extraidos.append(f"=== OBSERVAÇÕES / MEMORIAL / WHATSAPP ===\n{texto_livre.strip()}")

    for up in arquivos_up:
        raw = up.read()
        up.seek(0)
        fname = up.name.lower()
        if fname.endswith(".docx"):
            textos_extraidos.append(f"=== WORD: {up.name} ===\n{ler_docx_bytes(raw)}")
        elif fname.endswith((".xlsx", ".ods")):
            textos_extraidos.append(ler_planilha_sigef(raw, up.name))
        elif fname.endswith(".txt"):
            textos_extraidos.append(f"=== TXT: {up.name} ===\n{raw.decode('utf-8', errors='ignore')}")
        elif fname.endswith(".pdf"):
            partes_conteudo.append(types.Part.from_bytes(data=raw, mime_type="application/pdf"))
        elif fname.endswith((".png", ".jpg", ".jpeg", ".webp")):
            mime = "image/png" if fname.endswith(".png") else "image/jpeg"
            partes_conteudo.append(types.Part.from_bytes(data=raw, mime_type=mime))

    prompt = """
Você é um analista documental cartorial sênior de Minas Gerais (Facillyte Topografia).
Analise TODOS os documentos anexados (PDFs, fotos de RG/CPF/certidões/CCIR/ITR/CAR/matrículas, planilhas SIGEF e textos)
e retorne APENAS um JSON válido (sem markdown) com exatamente estas chaves:
{
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
  "glebas": [
    {"nome": "Gleba 01", "area_ha": 0.0, "perimetro_m": 0.0, "sigef": "", "valor_rs": ""}
  ],
  "confrontantes": [
    {
      "nome": "",
      "tipo": "PF",
      "doc": "",
      "rg": "",
      "qualificacao": "",
      "endereco": "",
      "matricula": "",
      "cns": "",
      "representante_ou_inventariante": "",
      "falecido": false
    }
  ]
}
Regras: não invente números ausentes (deixe vazio "" ou 0.0). Se detectar confrontante falecido ou espólio, marque "falecido": true e extraia o inventariante se constar.
"""
    if textos_extraidos:
        prompt += "\n\nCONTEÚDO EXTRAÍDO DE ARQUIVOS DE TEXTO/WORD/SIGEF:\n" + "\n\n".join(textos_extraidos)

    partes_conteudo.append(prompt)
    resp = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=partes_conteudo,
    )
    raw_text = resp.text.strip()
    raw_text = re.sub(r"^```json\s*", "", raw_text)
    raw_text = re.sub(r"\s*```$", "", raw_text)
    return json.loads(raw_text)


# ==============================================================================
# MOTOR DE AUDITORIA ANTI-DEVOLUTIVA (MG)
# ==============================================================================
def auditar_processo(d: dict) -> list[dict]:
    alertas = []  # {"nivel": "VERMELHO"|"AMARELO"|"VERDE", "regra": str, "detalhe": str}
    rito = d.get("rito", "")
    regra_rito = RITOS_MG.get(rito, {})

    # 1. Qualificação do proprietário e cônjuge
    if not d.get("proprietario_nome") or not d.get("proprietario_doc"):
        alertas.append({"nivel": "VERMELHO", "regra": "Qualificação do Proprietário", "detalhe": "Nome ou CPF/CNPJ do proprietário não informado (Provimento nº 61/2017/CNJ)."})
    else:
        alertas.append({"nivel": "VERDE", "regra": "Qualificação do Proprietário", "detalhe": f"{d['proprietario_nome']} ({d['proprietario_doc']}) identificado."})

    ec = (d.get("proprietario_estado_civil") or "").lower()
    if "casad" in ec or "união" in ec or "uniao" in ec:
        if not d.get("conjuge_nome") or not d.get("regime_bens"):
            alertas.append({"nivel": "VERMELHO", "regra": "Cônjuge e Regime de Bens", "detalhe": "Proprietário casado/união estável sem nome do cônjuge ou regime de bens informado (causa frequente de nota devolutiva em MG)."})
        else:
            alertas.append({"nivel": "VERDE", "regra": "Cônjuge e Regime de Bens", "detalhe": f"Cônjuge {d['conjuge_nome']} e regime '{d['regime_bens']}' preenchidos."})
            if "parcial" not in d.get("regime_bens", "").lower():
                alertas.append({"nivel": "AMARELO", "regra": "Pacto Antenupcial", "detalhe": "Regime diverso da comunhão parcial: conferir se o Pacto Antenupcial está registrado no Livro 3 (CRI)."})

    if d.get("proprietario_tipo") in ("PJ", "Espólio") and not d.get("representante_legal_ou_inventariante"):
        alertas.append({"nivel": "VERMELHO", "regra": "Representação Legal / Espólio", "detalhe": f"Proprietário é {d['proprietario_tipo']}, mas não há representante legal ou inventariante informado."})

    # 2. Valor Real e de Mercado (Lei Estadual 15.424/2004)
    if regra_rito.get("exige_valor_mercado"):
        if not d.get("valor_mercado_num"):
            alertas.append({"nivel": "VERMELHO", "regra": "Valor Real e de Mercado (Lei Est. 15.424/2004)", "detalhe": "Valor de mercado ausente — obrigatório para cálculo da Tabela 4 de emolumentos em MG (art. 10, § 3º, XVI)."})
        else:
            alertas.append({"nivel": "VERDE", "regra": "Valor Real e de Mercado (Lei Est. 15.424/2004)", "detalhe": f"Valor declarado: R$ {d['valor_mercado_num']}."})

    # 3. Conferência Cruzada de Áreas (Matrícula x CCIR x Memorial)
    a_mat = float(d.get("area_matricula_ha") or 0)
    a_ccir = float(d.get("area_ccir_ha") or 0)
    a_mem = float(d.get("area_memorial_ha") or 0)
    if a_mat > 0 and a_mem > 0:
        diff = abs(a_mat - a_mem)
        if diff > 0.0005:
            if rito == "Georreferenciamento (SIGEF/INCRA)":
                alertas.append({"nivel": "AMARELO", "regra": "Divergência Matrícula x Memorial", "detalhe": f"Área da matrícula ({a_mat:.4f} ha) diverge do memorial ({a_mem:.4f} ha). Sugestão automática: acumular Georreferenciamento com Retificação Administrativa (art. 891, § 9º, Prov. Conj. 93/2020)."})
            elif rito == "Inserção de Dados / Medidas Atualizadas":
                alertas.append({"nivel": "VERMELHO", "regra": "Alteração de Área na Inserção de Medidas", "detalhe": f"A área apurada ({a_mem:.4f} ha) diverge da matrícula ({a_mat:.4f} ha). Enquadrar como Retificação Administrativa (art. 213, II)."})
            else:
                alertas.append({"nivel": "AMARELO", "regra": "Diferença de Área Registrada x Apurada", "detalhe": f"Diferença de {diff:.4f} ha entre matrícula ({a_mat:.4f} ha) e memorial ({a_mem:.4f} ha) — conferir anuência de todos os confrontantes."})
        else:
            alertas.append({"nivel": "VERDE", "regra": "Compatibilidade Matrícula x Memorial", "detalhe": f"Área do memorial ({a_mem:.4f} ha) compatível com a matrícula."})

    if a_ccir > 0 and a_mem > 0 and abs(a_ccir - a_mem) > 0.05:
        alertas.append({"nivel": "AMARELO", "regra": "Divergência CCIR/ITR x Memorial", "detalhe": f"Área do CCIR ({a_ccir:.4f} ha) diverge da área do memorial ({a_mem:.4f} ha). Atualizar CCIR/SNCR se exigido pela serventia."})

    # 4. Soma de Áreas e Fração Mínima de Parcelamento (FMP)
    if regra_rito.get("exige_fmp"):
        fmp = float(d.get("fmp_municipio_ha") or 0)
        glebas = [g for g in d.get("glebas", []) if float(g.get("area_ha") or 0) > 0]
        if not glebas:
            alertas.append({"nivel": "VERMELHO", "regra": "Quadro de Glebas / Parcelas", "detalhe": "Nenhuma gleba/parcela informada para o desmembramento ou estremação."})
        else:
            soma = sum(float(g.get("area_ha") or 0) for g in glebas)
            alvo = a_mem if a_mem > 0 else a_mat
            if alvo > 0 and abs(soma - alvo) > 0.001:
                alertas.append({"nivel": "VERMELHO", "regra": "Fechamento de Áreas (Glebas + Remanescente)", "detalhe": f"Soma das parcelas ({soma:.4f} ha) não bate com a área total ({alvo:.4f} ha). Diferença: {abs(soma - alvo):.4f} ha."})
            else:
                alertas.append({"nivel": "VERDE", "regra": "Fechamento de Áreas", "detalhe": f"Soma das {len(glebas)} parcelas ({soma:.4f} ha) fecha com a área total."})
            if fmp > 0:
                abaixo = [g["nome"] for g in glebas if float(g.get("area_ha") or 0) < fmp]
                if abaixo:
                    alertas.append({"nivel": "VERMELHO", "regra": "Fração Mínima de Parcelamento (FMP)", "detalhe": f"Parcela(s) abaixo da FMP ({fmp:.4f} ha): {', '.join(abaixo)} — vedado pelo art. 65 da Lei nº 4.504/64 e arts. 991/1.149 do Prov. Conj. 93/2020."})
                else:
                    alertas.append({"nivel": "VERDE", "regra": "Fração Mínima de Parcelamento (FMP)", "detalhe": f"Todas as parcelas atendem à FMP de {fmp:.4f} ha."})

    # 5. Confrontantes (grafia, matrícula, falecido sem espólio, PJ sem representante)
    confs = [c for c in d.get("confrontantes", []) if (c.get("nome") or "").strip()]
    if regra_rito.get("exige_confrontantes") and not confs:
        alertas.append({"nivel": "VERMELHO", "regra": "Confrontantes", "detalhe": "Nenhum confrontante cadastrado para este rito."})
    mem_txt = (d.get("memorial_texto") or "").lower()
    for c in confs:
        nome = c.get("nome", "").strip()
        if not c.get("matricula"):
            alertas.append({"nivel": "AMARELO", "regra": f"Matrícula de Confrontante ({nome})", "detalhe": "Matrícula/Transcrição/Posse do confrontante não preenchida."})
        if c.get("falecido") or "espólio" in nome.lower() or "espolio" in nome.lower():
            if not c.get("representante_ou_inventariante"):
                alertas.append({"nivel": "VERMELHO", "regra": f"Confrontante Espólio ({nome})", "detalhe": "Confrontante falecido/espólio sem indicação de Inventariante (exige termo de inventariante + certidão de óbito)."})
        if c.get("tipo") == "PJ" and not c.get("representante_ou_inventariante"):
            alertas.append({"nivel": "VERMELHO", "regra": f"Confrontante PJ ({nome})", "detalhe": "Pessoa Jurídica sem representante legal informado (juntar contrato social/ata)."})
        if mem_txt and nome.lower() not in mem_txt:
            alertas.append({"nivel": "AMARELO", "regra": f"Grafia no Memorial ({nome})", "detalhe": f"O nome exato '{nome}' não foi encontrado no texto do memorial descritivo — confira se há divergência de grafia."})

    # 6. Cadastros e ART
    faltantes = []
    for k, lbl in [("art_numero", "ART"), ("incra_sncr", "CCIR/SNCR"), ("car_numero", "CAR"), ("comarca_uf", "Comarca do CRI"), ("matricula_numero", "Matrícula")]:
        if not (d.get(k) or "").strip():
            faltantes.append(lbl)
    if faltantes:
        alertas.append({"nivel": "AMARELO", "regra": "Campos Cadastrais Pendentes", "detalhe": f"Preencher antes do protocolo: {', '.join(faltantes)}."})
    else:
        alertas.append({"nivel": "VERDE", "regra": "Cadastros e ART", "detalhe": "ART, CCIR/SNCR, CAR, Matrícula e Comarca preenchidos."})

    return alertas


# ==============================================================================
# GERADOR DE DOCUMENTOS NO TIMBRADO OFICIAL FACILLYTE (.DOCX + .PDF)
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


def set_cell_mar(cell, top=55, bottom=55, left=80, right=80):
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

        # Cabeçalho 3 colunas (modelo oficial Facillyte Engenharia LTDA)
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
            r.font.bold = True
            r.font.color.rgb = RGB_PRIMARY

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
            rr.font.name = "Calibri"
            rr.font.size = Pt(sz)
            rr.font.bold = bld
            rr.font.color.rgb = col

        p2 = htbl.rows[0].cells[2].paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if os.path.exists(LOGO_TOPO):
            p2.add_run().add_picture(LOGO_TOPO, width=Cm(2.15))
        else:
            r = p2.add_run("FACILLYTE\nTOPOGRAFIA")
            r.font.bold = True
            r.font.color.rgb = RGB_ACCENT

        p_acc = hdr.add_paragraph()
        p_acc.paragraph_format.space_before = Pt(1)
        p_acc.paragraph_format.space_after = Pt(3)
        p_acc.paragraph_format.line_spacing = Pt(2)
        p_acc._p.get_or_add_pPr().append(
            parse_xml(f'<w:pBdr {nsdecls("w")}><w:top w:val="single" w:sz="8" w:space="0" w:color="{HEX_ACCENT}"/></w:pBdr>')
        )

        # Rodapé
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
    r.font.name = "Calibri"
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    return p


def add_assinatura(doc, nome, legenda):
    p1 = add_p(doc, "____________________________________________________________", center=True, space_after=2)
    p1.paragraph_format.space_before = Pt(12)
    p1.paragraph_format.keep_with_next = True
    p2 = add_p(doc, nome, bold=True, center=True, space_after=1)
    p2.paragraph_format.keep_with_next = True
    add_p(doc, legenda, center=True, size=10, color=RGB_MUTED, space_after=6)


def qualificacao_proprietario(d: dict) -> str:
    partes = [d.get("proprietario_nome", "[PROPRIETÁRIO]")]
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

    docs_map = {}

    # 1. Requerimento Principal
    doc_req = novo_doc_timbrado()
    add_p(doc_req, f"ILUSTRÍSSIMO(A) SENHOR(A) OFICIAL(A) DO CARTÓRIO DE REGISTRO DE IMÓVEIS DA COMARCA DE {comarca.upper()}", bold=True, justify=False, space_after=8)
    add_p(doc_req, f"REQUERIMENTO DE {rito.upper()}", bold=True, center=True, size=13, color=RGB_PRIMARY, space_after=3)
    add_p(doc_req, fundamento, center=True, size=10, color=RGB_MUTED, space_after=8)
    add_p(doc_req, f"{qual_prop}, juntamente com o Responsável Técnico MARCOS FRANCIEL FERNANDES DOS SANTOS, Engenheiro Civil, CREA-MG 248849/D, RNP 141908491-7/MG, INCRA YIDB, com escritório à Rua Ayrton Senna, 425, Vila Nova, Pai Pedro/MG, vêm requerer o processamento e averbação de {rito} no imóvel denominado {d.get('imovel_denominacao', '')}, Município de {mun}/MG, Matrícula nº {d.get('matricula_numero', '')}, Livro {d.get('livro_numero', '2')}, CNS {d.get('cartorio_cns', '')}, INCRA/SNCR {d.get('incra_sncr', '')}, CIB/NIRF {d.get('cib_nirf', '')}, CAR {d.get('car_numero', '')}, Certificação SIGEF {d.get('sigef_codigo', 'N/A')}.")
    add_p(doc_req, f"Área Registrada: {d.get('area_matricula_ha', 0):.4f} ha | Área Apurada/Memorial: {d.get('area_memorial_ha', 0):.4f} ha | Perímetro: {d.get('perimetro_m', 0):.3f} m | Valor Real e de Mercado (Lei Est. 15.424/2004, Tabela 4): R$ {d.get('valor_mercado_num', '')} ({d.get('valor_mercado_extenso', '')}).", bold=True)
    add_p(doc_req, "Os requerentes e o Responsável Técnico declaram, sob responsabilidade civil e criminal (art. 213, § 14, da Lei nº 6.015/1973), que o levantamento foi realizado in loco (Datum SIRGAS 2000), respeitando integralmente as divisas existentes, faixas de domínio público e os direitos de todos os confrontantes, inexistindo invasão ou ação judicial conexa.")
    add_p(doc_req, f"Nestes termos, pede deferimento. {mun}/MG, {data_hoje}.", justify=False, space_after=10)
    add_assinatura(doc_req, d.get("proprietario_nome", "[PROPRIETÁRIO]"), f"Requerente — CPF/CNPJ: {d.get('proprietario_doc', '')}")
    if d.get("conjuge_nome"):
        add_assinatura(doc_req, d["conjuge_nome"], f"Cônjuge — CPF: {d.get('conjuge_cpf', '')}")
    add_assinatura(doc_req, "MARCOS FRANCIEL FERNANDES DOS SANTOS", rt_legenda)
    docs_map["01_Requerimento_CRI.docx"] = doc_req

    # 2. Declarações Individuais de Anuência dos Confrontantes
    for idx, c in enumerate(d.get("confrontantes", []), start=1):
        if not (c.get("nome") or "").strip():
            continue
        doc_c = novo_doc_timbrado()
        add_p(doc_c, "DECLARAÇÃO DE ANUÊNCIA E RECONHECIMENTO DE LIMITES (CONFRONTANTE)", bold=True, center=True, size=13, color=RGB_PRIMARY, space_after=3)
        add_p(doc_c, "Art. 213, II e § 10, da Lei nº 6.015/1973 • Arts. 891 e 898 do Provimento Conjunto nº 93/2020 (CGJ/MG)", center=True, size=10, color=RGB_MUTED, space_after=8)
        rep = f", neste ato representado(a) por {c['representante_ou_inventariante']}" if c.get("representante_ou_inventariante") else ""
        add_p(doc_c, f"Eu, {c['nome']}{rep}, {c.get('qualificacao', '')}, RG {c.get('rg', '')}, CPF/CNPJ {c.get('doc', '')}, residente/sediado(a) em {c.get('endereco', '')}, titular/ocupante do imóvel confrontante Matrícula/Posse {c.get('matricula', '')} (CNS {c.get('cns', '')}), DECLARO, sob as penas da lei (art. 299 do Código Penal), que concordo integralmente com as medidas perimetrais, coordenadas, azimutes e divisas comuns constantes da planta e memorial descritivo do imóvel {d.get('imovel_denominacao', '')}, Matrícula nº {d.get('matricula_numero', '')} (CRI {comarca}), de propriedade de {d.get('proprietario_nome', '')}, levantado pelo Eng. Civil Marcos Franciel Fernandes dos Santos (CREA-MG 248849/D, INCRA YIDB, ART nº {d.get('art_numero', '')}).")
        add_p(doc_c, "Declaro ciência de que, nos termos do art. 213, § 10, da Lei nº 6.015/1973 e art. 898 do Provimento Conjunto nº 93/2020/CGJ-MG, minha anuência supre a participação do cônjuge e de eventuais outros condôminos titulares do imóvel limítrofe.")
        add_p(doc_c, f"{mun}/MG, {data_hoje}.", justify=False, space_after=10)
        safe_name = re.sub(r"[^A-Za-z0-9]+", "_", c["nome"])[:30]
        add_assinatura(doc_c, c.get("representante_ou_inventariante") or c["nome"], f"Confrontante ({c['nome']}) — CPF/CNPJ: {c.get('doc', '')}")
        docs_map[f"02_Anuencia_{idx:02d}_{safe_name}.docx"] = doc_c

    # 3. Declaração Conjunta do Proprietário e RT
    doc_dc = novo_doc_timbrado()
    add_p(doc_dc, "DECLARAÇÃO CONJUNTA DO(S) PROPRIETÁRIO(S) E DO RESPONSÁVEL TÉCNICO", bold=True, center=True, size=13, color=RGB_PRIMARY, space_after=3)
    add_p(doc_dc, "Art. 213, § 14, da Lei nº 6.015/1973 • Provimento Conjunto nº 93/2020 (CGJ/MG)", center=True, size=10, color=RGB_MUTED, space_after=8)
    add_p(doc_dc, f"Referência: Imóvel {d.get('imovel_denominacao', '')}, {mun}/MG, Matrícula nº {d.get('matricula_numero', '')}, CRI de {comarca}.", bold=True)
    add_p(doc_dc, f"Responsável Técnico: Eu, MARCOS FRANCIEL FERNANDES DOS SANTOS, Engenheiro Civil, CREA-MG 248849/D, RNP 141908491-7/MG, INCRA YIDB (ART nº {d.get('art_numero', '')}), ATESTO sob as penas da lei que efetuei pessoalmente, in loco, o levantamento georreferenciado (SIRGAS 2000), sendo exatos os dados da planta e memorial descritivo e plenamente respeitadas as divisas particulares e faixas de domínio público.")
    add_p(doc_dc, f"Proprietário(s): {qual_prop}, ATESTAMOS sob as penas da lei serem verdadeiras todas as informações apresentadas, sem alteração de divisas nem invasão de áreas vizinhas, com plena ciência do art. 213, § 14, da Lei nº 6.015/1973 e do art. 299 do Código Penal.")
    add_p(doc_dc, f"{mun}/MG, {data_hoje}.", justify=False, space_after=10)
    add_assinatura(doc_dc, "MARCOS FRANCIEL FERNANDES DOS SANTOS", rt_legenda)
    add_assinatura(doc_dc, d.get("proprietario_nome", "[PROPRIETÁRIO]"), f"Proprietário(a) — CPF/CNPJ: {d.get('proprietario_doc', '')}")
    docs_map["03_Declaracao_Conjunta_Proprietario_RT.docx"] = doc_dc

    # 4. Declaração de Valor Real e de Mercado (MG)
    doc_vm = novo_doc_timbrado()
    add_p(doc_vm, f"ILUSTRÍSSIMO(A) SENHOR(A) OFICIAL(A) DO CARTÓRIO DE REGISTRO DE IMÓVEIS DA COMARCA DE {comarca.upper()}", bold=True, justify=False, space_after=8)
    add_p(doc_vm, "DECLARAÇÃO DE VALOR REAL E DE MERCADO DO IMÓVEL", bold=True, center=True, size=13, color=RGB_PRIMARY, space_after=3)
    add_p(doc_vm, "Art. 10, § 3º, XVI, Lei Estadual nº 15.424/2004 (Tabela 4) • Arts. 134 a 138 do Prov. Conj. nº 93/2020 (CGJ/MG)", center=True, size=10, color=RGB_MUTED, space_after=8)
    add_p(doc_vm, f"{qual_prop}, proprietário(a) do imóvel {d.get('imovel_denominacao', '')}, Matrícula nº {d.get('matricula_numero', '')}, CRI de {comarca}, requerente de {rito}, em atendimento aos arts. 134 a 138 do Provimento Conjunto nº 93/2020/CGJ-MG e ao art. 10, § 3º, XVI, da Lei Estadual nº 15.424/2004 (alterada pelas Leis nº 25.125/2024 e nº 25.367/2025), DECLARA, sob responsabilidade civil e penal, que o valor real e de mercado do imóvel é de R$ {d.get('valor_mercado_num', '')} ({d.get('valor_mercado_extenso', '')}).")
    add_p(doc_vm, "Declara ciência de que o valor destina-se ao recolhimento de emolumentos e TFJ, sujeito a arbitramento judicial se manifestamente inferior ao real (arts. 134 a 138 do Prov. Conj. 93/2020), bem como plena ciência do art. 299 do Código Penal Brasileiro.")
    add_p(doc_vm, f"{mun}/MG, {data_hoje}.", justify=False, space_after=10)
    add_assinatura(doc_vm, d.get("proprietario_nome", "[PROPRIETÁRIO]"), f"Proprietário(a) — CPF/CNPJ: {d.get('proprietario_doc', '')}")
    docs_map["04_Declaracao_Valor_Real_Mercado_MG.docx"] = doc_vm

    # 5. Memorial Descritivo Padronizado
    doc_mem = novo_doc_timbrado()
    add_p(doc_mem, "MEMORIAL DESCRITIVO", bold=True, center=True, size=13.5, color=RGB_PRIMARY, space_after=4)
    add_p(doc_mem, f"Imóvel: {d.get('imovel_denominacao', '')} • Matrícula: {d.get('matricula_numero', '')} • Município: {mun}/MG • Comarca: {comarca} (CNS {d.get('cartorio_cns', '')})", bold=True, justify=False, space_after=4)
    add_p(doc_mem, f"Proprietário(a): {d.get('proprietario_nome', '')} ({d.get('proprietario_doc', '')}) • INCRA: {d.get('incra_sncr', '')} • Área: {d.get('area_memorial_ha', 0):.4f} ha • Perímetro: {d.get('perimetro_m', 0):.3f} m • Datum: SIRGAS 2000", justify=False, space_after=8)
    add_p(doc_mem, d.get("memorial_texto") or "[Inserir/colar descrição completa do perímetro com coordenadas, azimutes, distâncias e confrontantes]")
    add_p(doc_mem, f"{mun}/MG, {data_hoje}.", justify=False, space_after=8)
    add_assinatura(doc_mem, "MARCOS FRANCIEL FERNANDES DOS SANTOS", rt_legenda)
    add_assinatura(doc_mem, d.get("proprietario_nome", "[PROPRIETÁRIO]"), f"Proprietário(a) — CPF/CNPJ: {d.get('proprietario_doc', '')}")
    docs_map["05_Memorial_Descritivo_Timbrado.docx"] = doc_mem

    # 6. Checklist Final Anti-Devolutiva + Resultado da Auditoria
    doc_chk = novo_doc_timbrado()
    add_p(doc_chk, "CHECKLIST FINAL E RELATÓRIO ANTI-DEVOLUTIVA (MG)", bold=True, center=True, size=13, color=RGB_PRIMARY, space_after=6)
    add_p(doc_chk, f"Processo: {rito} — {d.get('imovel_denominacao', '')} (Matrícula {d.get('matricula_numero', '')} — {comarca})", bold=True, justify=False)
    for a in alertas:
        ico = {"VERDE": "[OK]", "AMARELO": "[ATENÇÃO]", "VERMELHO": "[BLOQUEIO]"}.get(a["nivel"], "")
        add_p(doc_chk, f"{ico} {a['regra']}: {a['detalhe']}", justify=False, space_after=4)
    add_p(doc_chk, "Conferência Física/Digital Pré-Protocolo: [ ] ART quitada  [ ] CCIR quitado  [ ] CND ITR  [ ] Recibo CAR  [ ] Planta e Memorial A4  [ ] Firmas reconhecidas ou PDF/A (Prov. 149/2023 CNJ).", bold=True, justify=False, space_after=6)
    docs_map["06_Checklist_Auditoria_AntiDevolutiva.docx"] = doc_chk

    # Empacotar ZIP (DOCX + PDF opcional se LibreOffice instalado)
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
# INTERFACE STREAMLIT (4 ETAPAS LINEARES)
# ==============================================================================
init_db()

with st.sidebar:
    if os.path.exists(LOGO_TOPO):
        st.image(LOGO_TOPO, width=130)
    st.title("Facillyte Docs v2")
    st.caption("Funcionário Virtual Cartorial — MG")
    api_key = st.secrets.get("GEMINI_API_KEY", "") if hasattr(st, "secrets") else ""
    api_key_input = st.text_input("Chave API Gemini (se não estiver nos Secrets)", value=api_key, type="password")
    st.divider()
    st.subheader("Histórico de Processos (SQLite)")
    hist = listar_processos()
    if hist:
        opcoes = {f"#{r[0]} • {r[3]} ({r[5]}) — {r[1]}": r[7] for r in hist}
        sel = st.selectbox("Carregar processo salvo", ["(Novo processo)"] + list(opcoes.keys()))
        if sel != "(Novo processo)" and st.button("Carregar dados"):
            st.session_state["dados"] = json.loads(opcoes[sel])
            st.success("Processo carregado!")
    else:
        st.caption("Nenhum processo salvo ainda.")

if "dados" not in st.session_state:
    st.session_state["dados"] = {
        "rito": list(RITOS_MG.keys())[0],
        "proprietario_tipo": "PF",
        "proprietario_nac": "brasileiro(a)",
        "livro_numero": "2",
        "fmp_municipio_ha": 3.0,
        "glebas": [{"nome": "Gleba 01", "area_ha": 0.0, "perimetro_m": 0.0, "sigef": "", "valor_rs": ""}],
        "confrontantes": [{"nome": "", "tipo": "PF", "doc": "", "rg": "", "qualificacao": "", "endereco": "", "matricula": "", "cns": "", "representante_ou_inventariante": "", "falecido": False}],
    }

dados = st.session_state["dados"]

st.header("📐 Funcionário Virtual — Peças Técnicas Cartoriais (MG)")
tab1, tab2, tab3, tab4 = st.tabs([
    "1️⃣ Ingestão Multimodal",
    "2️⃣ Escolha do Rito & Revisão",
    "3️⃣ Auditoria Anti-Devolutiva",
    "4️⃣ Gerar Pacote no Timbrado (.zip)",
])

# ── ETAPA 1 ──
with tab1:
    st.subheader("Envie todos os documentos brutos do cliente")
    arquivos = st.file_uploader(
        "PDFs, fotos/imagens (RG, CPF, certidões, matrículas, CCIR, ITR, CAR), Word (.docx) e planilhas SIGEF (.xlsx / .ods)",
        type=["pdf", "png", "jpg", "jpeg", "webp", "docx", "xlsx", "ods", "txt"],
        accept_multiple_files=True,
    )
    texto_livre = st.text_area(
        "Texto livre (memorial do TopoCAD/AutoCAD, anotações ou dados colados do WhatsApp do cliente)",
        height=160,
    )
    if st.button("🤖 Extrair dados automaticamente com IA", type="primary"):
        if not api_key_input:
            st.error("Informe a chave GEMINI_API_KEY na barra lateral ou nos Secrets do Streamlit Cloud.")
        elif not arquivos and not texto_livre.strip():
            st.warning("Envie ao menos um arquivo ou cole um texto para extrair.")
        else:
            with st.spinner("Lendo documentos, planilhas SIGEF e imagens..."):
                try:
                    extraido = extrair_com_gemini(api_key_input, arquivos, texto_livre)
                    extraido["rito"] = dados.get("rito", list(RITOS_MG.keys())[0])
                    st.session_state["dados"] = extraido
                    st.success("Extração concluída! Avance para a aba 2 para revisar.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro na extração: {e}")

# ── ETAPA 2 ──
with tab2:
    st.subheader("Rito Cartorial (Minas Gerais) e Ficha Estruturada")
    dados["rito"] = st.selectbox("Selecione o Rito (MG)", list(RITOS_MG.keys()), index=list(RITOS_MG.keys()).index(dados.get("rito", list(RITOS_MG.keys())[0])))
    st.info(f"**Base Legal Aplicada:** {RITOS_MG[dados['rito']]['fundamento']}")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Proprietário / Requerente**")
        dados["proprietario_tipo"] = st.selectbox("Tipo", ["PF", "PJ", "Espólio"], index=["PF", "PJ", "Espólio"].index(dados.get("proprietario_tipo", "PF")))
        dados["proprietario_nome"] = st.text_input("Nome Completo / Razão Social", value=dados.get("proprietario_nome", ""))
        dados["proprietario_doc"] = st.text_input("CPF ou CNPJ", value=dados.get("proprietario_doc", ""))
        dados["proprietario_rg"] = st.text_input("RG / Órgão Expedidor", value=dados.get("proprietario_rg", ""))
        dados["proprietario_estado_civil"] = st.text_input("Estado Civil", value=dados.get("proprietario_estado_civil", ""))
        dados["regime_bens"] = st.text_input("Regime de Bens", value=dados.get("regime_bens", ""))
        dados["proprietario_prof"] = st.text_input("Profissão", value=dados.get("proprietario_prof", ""))
        dados["proprietario_endereco"] = st.text_input("Endereço Completo", value=dados.get("proprietario_endereco", ""))
        dados["conjuge_nome"] = st.text_input("Cônjuge (se aplicável)", value=dados.get("conjuge_nome", ""))
        dados["conjuge_cpf"] = st.text_input("CPF do Cônjuge", value=dados.get("conjuge_cpf", ""))
        dados["representante_legal_ou_inventariante"] = st.text_input("Representante Legal / Inventariante (PJ ou Espólio)", value=dados.get("representante_legal_ou_inventariante", ""))

    with col2:
        st.markdown("**Imóvel, Cadastros e Áreas**")
        dados["imovel_denominacao"] = st.text_input("Denominação do Imóvel", value=dados.get("imovel_denominacao", ""))
        dados["imovel_municipio"] = st.text_input("Município (MG)", value=dados.get("imovel_municipio", ""))
        dados["comarca_uf"] = st.text_input("Comarca do CRI (ex.: Porteirinha/MG)", value=dados.get("comarca_uf", ""))
        dados["cartorio_cns"] = st.text_input("CNS do Cartório", value=dados.get("cartorio_cns", ""))
        dados["matricula_numero"] = st.text_input("Nº da Matrícula / Transcrição", value=dados.get("matricula_numero", ""))
        dados["incra_sncr"] = st.text_input("Código INCRA / SNCR", value=dados.get("incra_sncr", ""))
        dados["cib_nirf"] = st.text_input("CIB / NIRF", value=dados.get("cib_nirf", ""))
        dados["car_numero"] = st.text_input("Recibo CAR", value=dados.get("car_numero", ""))
        dados["sigef_codigo"] = st.text_input("Código de Certificação SIGEF", value=dados.get("sigef_codigo", ""))
        dados["art_numero"] = st.text_input("Nº da ART (CREA-MG)", value=dados.get("art_numero", ""))
        c_a1, c_a2, c_a3 = st.columns(3)
        dados["area_matricula_ha"] = c_a1.number_input("Área Matrícula (ha)", value=float(dados.get("area_matricula_ha") or 0.0), format="%.4f")
        dados["area_ccir_ha"] = c_a2.number_input("Área CCIR (ha)", value=float(dados.get("area_ccir_ha") or 0.0), format="%.4f")
        dados["area_memorial_ha"] = c_a3.number_input("Área Memorial (ha)", value=float(dados.get("area_memorial_ha") or 0.0), format="%.4f")
        c_p1, c_p2 = st.columns(2)
        dados["perimetro_m"] = c_p1.number_input("Perímetro (m)", value=float(dados.get("perimetro_m") or 0.0), format="%.3f")
        dados["fmp_municipio_ha"] = c_p2.number_input("FMP do Município (ha)", value=float(dados.get("fmp_municipio_ha") or 3.0), format="%.4f")
        c_v1, c_v2 = st.columns(2)
        dados["valor_mercado_num"] = c_v1.text_input("Valor de Mercado (R$)", value=dados.get("valor_mercado_num", ""))
        dados["valor_mercado_extenso"] = c_v2.text_input("Valor por Extenso", value=dados.get("valor_mercado_extenso", ""))

    st.markdown("**Memorial Descritivo (texto completo do perímetro)**")
    dados["memorial_texto"] = st.text_area("Memorial", value=dados.get("memorial_texto", ""), height=130, label_visibility="collapsed")

    if RITOS_MG[dados["rito"]]["exige_fmp"]:
        st.markdown("**Glebas / Parcelas e Remanescente (Desmembramento / Estremação)**")
        df_glebas = pd.DataFrame(dados.get("glebas") or [{"nome": "Gleba 01", "area_ha": 0.0, "perimetro_m": 0.0, "sigef": "", "valor_rs": ""}])
        edited_glebas = st.data_editor(df_glebas, num_rows="dynamic", use_container_width=True)
        dados["glebas"] = edited_glebas.to_dict(orient="records")

    st.markdown("**Lista de Confrontantes**")
    df_confs = pd.DataFrame(
        dados.get("confrontantes")
        or [{"nome": "", "tipo": "PF", "doc": "", "rg": "", "qualificacao": "", "endereco": "", "matricula": "", "cns": "", "representante_ou_inventariante": "", "falecido": False}]
    )
    edited_confs = st.data_editor(df_confs, num_rows="dynamic", use_container_width=True)
    dados["confrontantes"] = edited_confs.to_dict(orient="records")
    st.session_state["dados"] = dados

# ── ETAPA 3 ──
with tab3:
    st.subheader("Painel de Auditoria Anti-Devolutiva (CGJ/MG)")
    alertas = auditar_processo(dados)
    n_verm = sum(1 for a in alertas if a["nivel"] == "VERMELHO")
    n_amar = sum(1 for a in alertas if a["nivel"] == "AMARELO")
    n_verd = sum(1 for a in alertas if a["nivel"] == "VERDE")
    m1, m2, m3 = st.columns(3)
    m1.metric("🔴 Bloqueios Críticos", n_verm)
    m2.metric("🟡 Alertas de Atenção", n_amar)
    m3.metric("🟢 Itens Validados", n_verd)

    for a in alertas:
        if a["nivel"] == "VERMELHO":
            st.error(f"**{a['regra']}:** {a['detalhe']}")
        elif a["nivel"] == "AMARELO":
            st.warning(f"**{a['regra']}:** {a['detalhe']}")
        else:
            st.success(f"**{a['regra']}:** {a['detalhe']}")

# ── ETAPA 4 ──
with tab4:
    st.subheader("Geração do Pacote Completo no Papel Timbrado Oficial")
    alertas_finais = auditar_processo(dados)
    if any(a["nivel"] == "VERMELHO" for a in alertas_finais):
        st.warning("Existem alertas vermelhos na aba 3. Você ainda pode gerar o pacote para rascunho, mas recomenda-se corrigir antes do protocolo.")
    gerar_pdf = st.checkbox("Incluir versão PDF pronta para impressão além do Word (.docx)", value=True)
    if st.button("📦 Salvar no Banco e Gerar Pacote (.zip)", type="primary"):
        salvar_processo(dados)
        zip_bytes = gerar_pacote_zip(dados, alertas_finais, gerar_pdf=gerar_pdf)
        nome_arq = re.sub(r"[^A-Za-z0-9]+", "_", f"Dossie_{dados.get('imovel_denominacao') or 'Imovel'}_Mat_{dados.get('matricula_numero') or 'SN'}") + ".zip"
        st.success("Pacote gerado e processo salvo no histórico!")
        st.download_button(
            label=f"⬇️ Baixar {nome_arq}",
            data=zip_bytes,
            file_name=nome_arq,
            mime="application/zip",
        )
