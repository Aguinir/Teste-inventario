"""
Inventário de Frota — Automação de Evidências
OCR via Tesseract (pytesseract) — sem IA, sem API externa.
Deploy: Streamlit Cloud (GitHub)
"""

import streamlit as st
import fitz                     # PyMuPDF
from PIL import Image, ImageOps, ImageFilter, ImageEnhance
import pytesseract
import io, zipfile, re, numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment

# ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Inventário de Frota",
    page_icon="🚛",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
#MainMenu, header, footer,
[data-testid="stToolbar"],
[data-testid="stDecoration"]   { display: none !important; }

html, body, [class*="css"]     { font-family: 'Segoe UI', Arial, sans-serif; }
.main .block-container         { padding-top: 1.8rem; max-width: 1120px; }

.app-header {
    background: linear-gradient(135deg, #0d2137 0%, #1565C0 100%);
    border-radius: 12px;
    padding: 1.4rem 2rem;
    color: #fff;
    margin-bottom: 1.5rem;
}
.app-header h1 { margin: 0; font-size: 1.55rem; font-weight: 700; }
.app-header p  { margin: .3rem 0 0; opacity: .8; font-size: .88rem; }

.upload-card {
    background: #fff;
    border: 1px solid #dce3ef;
    border-radius: 10px;
    padding: 1.1rem 1.3rem .5rem;
}
.upload-card h3 { margin: 0 0 .25rem; color: #1565C0; font-size: .95rem; }
.upload-card p  { color: #666; font-size: .8rem; margin: 0 0 .6rem; }

.metrics       { display: flex; gap: .9rem; margin: 1rem 0; }
.metric-box {
    flex: 1; background: #fff;
    border: 1px solid #dce3ef;
    border-radius: 10px;
    padding: 1rem; text-align: center;
}
.metric-box .val  { font-size: 2rem; font-weight: 700; color: #1565C0; line-height: 1.1; }
.metric-box .lbl  { font-size: .73rem; color: #777; text-transform: uppercase;
                    letter-spacing: .04em; margin-top: .3rem; }
.metric-box.green  .val { color: #2e7d32; }
.metric-box.orange .val { color: #e65100; }

hr { border: none; border-top: 1px solid #e8eaf0; margin: 1.2rem 0; }

div[data-testid="stButton"] > button[kind="primary"] {
    background: #1565C0; border: none;
    font-weight: 600; font-size: .97rem;
    height: 3rem; border-radius: 8px;
}
div[data-testid="stButton"] > button[kind="primary"]:hover { background: #0d47a1; }
div[data-testid="stButton"] > button:disabled { opacity: .45; }
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────
# Regex – placas brasileiras
# ──────────────────────────────────────────────────────────────
PAT_MERCOSUL = re.compile(r"[A-Z]{3}[0-9][A-Z][0-9]{2}")
PAT_ANTIGO   = re.compile(r"[A-Z]{3}[0-9]{4}")

TESS_CFG = r"--oem 3 --psm 11 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"


def normalizar(texto: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(texto).upper())


# ──────────────────────────────────────────────────────────────
# Pré-processamento da imagem para OCR
# ──────────────────────────────────────────────────────────────
def preprocessar(pil_img: Image.Image) -> list[Image.Image]:
    """
    Retorna múltiplas variantes da imagem para maximizar a taxa de acerto
    do Tesseract em placas com iluminação e ângulo variados.
    """
    variantes = []

    # Garante tamanho mínimo razoável
    w, h = pil_img.size
    if max(w, h) < 400:
        scale = 400 / max(w, h)
        pil_img = pil_img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)

    gray = pil_img.convert("L")

    # Variante 1: grayscale puro
    variantes.append(gray)

    # Variante 2: contraste reforçado
    enhanced = ImageEnhance.Contrast(gray).enhance(2.5)
    variantes.append(enhanced)

    # Variante 3: nitidez
    sharpened = ImageEnhance.Sharpness(enhanced).enhance(2.0)
    variantes.append(sharpened)

    # Variante 4: threshold OTSU via numpy
    arr = np.array(gray)
    threshold = arr.mean()
    binary = Image.fromarray(np.where(arr > threshold, 255, 0).astype(np.uint8))
    variantes.append(binary)

    # Variante 5: inversão (placas escuras com fundo claro)
    variantes.append(ImageOps.invert(binary))

    return variantes


# ──────────────────────────────────────────────────────────────
# OCR + extração de candidatos
# ──────────────────────────────────────────────────────────────
def ocr_candidatos(pil_img: Image.Image) -> tuple[set, str]:
    """
    Executa Tesseract em múltiplas variantes e coleta todos os candidatos.
    Retorna (set de candidatos normalizados, texto concatenado total).
    """
    candidatos = set()
    textos_totais = []

    for variante in preprocessar(pil_img):
        try:
            raw = pytesseract.image_to_string(variante, config=TESS_CFG)
        except Exception:
            continue

        norm = normalizar(raw)
        textos_totais.append(norm)

        # Busca de placas em janelas deslizantes de 6 a 10 chars
        for tamanho in (7, 6, 8, 9, 10):
            for inicio in range(len(norm) - tamanho + 1):
                janela = norm[inicio: inicio + tamanho]
                for pat in (PAT_MERCOSUL, PAT_ANTIGO):
                    if pat.fullmatch(janela):
                        candidatos.add(janela)

        # Sequências longas (chassi / VIN ≥ 10 chars)
        for parte in re.findall(r"[A-Z0-9]{10,}", norm):
            candidatos.add(parte)

    return candidatos, "".join(textos_totais)


# ──────────────────────────────────────────────────────────────
# Lookup da planilha
# ──────────────────────────────────────────────────────────────
def construir_lookup(ws):
    placa_lkp, chassis_lkp = {}, {}
    for linha in range(4, ws.max_row + 1):
        a = ws.cell(linha, 1).value
        b = ws.cell(linha, 2).value
        d = ws.cell(linha, 4).value
        for v in (a, b):
            if v:
                n = normalizar(str(v))
                if n:
                    placa_lkp[n] = linha
        if d:
            n = normalizar(str(d))
            if len(n) >= 5:
                chassis_lkp[n] = linha
    return placa_lkp, chassis_lkp


def buscar_na_planilha(candidatos, concat_total, placa_lkp, chassis_lkp):
    # 1. Placas exatas
    for c in candidatos:
        if c in placa_lkp:
            return c, placa_lkp[c]
    # 2. Chassi exato
    for c in candidatos:
        if c in chassis_lkp:
            return c, chassis_lkp[c]
    # 3. Chassi substring no texto total
    for chassi, linha in chassis_lkp.items():
        if len(chassi) >= 10 and chassi in concat_total:
            return chassi, linha
    return None, None


# ──────────────────────────────────────────────────────────────
# Extração de imagens dos PDFs
# ──────────────────────────────────────────────────────────────
def extrair_imagens(pdf_files) -> list[dict]:
    imagens = []
    for pdf_file in pdf_files:
        pdf_file.seek(0)
        try:
            doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
        except Exception:
            continue
        for pag_idx in range(len(doc)):
            pagina = doc[pag_idx]
            lista  = pagina.get_images(full=True)
            adicionadas = 0
            for img_info in lista:
                try:
                    base = doc.extract_image(img_info[0])
                    pil  = Image.open(io.BytesIO(base["image"])).convert("RGB")
                    w, h = pil.size
                    if w < 100 or h < 100:
                        continue
                    raw = io.BytesIO()
                    pil.save(raw, format="JPEG", quality=88)
                    imagens.append({
                        "pil": pil, "raw": raw.getvalue(),
                        "ext": "jpg", "src": pdf_file.name, "pag": pag_idx + 1,
                    })
                    adicionadas += 1
                except Exception:
                    pass
            if adicionadas == 0:
                try:
                    pix = pagina.get_pixmap(matrix=fitz.Matrix(1.8, 1.8))
                    pil = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    raw = io.BytesIO()
                    pil.save(raw, format="JPEG", quality=88)
                    imagens.append({
                        "pil": pil, "raw": raw.getvalue(),
                        "ext": "jpg", "src": pdf_file.name, "pag": pag_idx + 1,
                    })
                except Exception:
                    pass
    return imagens


# ──────────────────────────────────────────────────────────────
# Atualização do XLSX — apenas coluna M
# ──────────────────────────────────────────────────────────────
def atualizar_xlsx(xlsx_bytes: bytes, linhas_encontradas: set) -> bytes:
    wb = load_workbook(io.BytesIO(xlsx_bytes))
    ws = wb["Inventário"]
    aln = Alignment(horizontal="center")
    for linha in range(4, ws.max_row + 1):
        if not ws.cell(linha, 1).value:
            continue
        cel = ws.cell(linha, 13)
        if linha in linhas_encontradas:
            cel.value     = "SIM"
            cel.alignment = aln
        else:
            cel.value = None
    saida = io.BytesIO()
    wb.save(saida)
    return saida.getvalue()


def montar_zip(imagens: list) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, img in enumerate(imagens):
            nome = (
                img["src"].replace(".pdf", "").replace(" ", "_")
                + f"_pag{img['pag']}_{i:04d}.{img['ext']}"
            )
            zf.writestr(nome, img["raw"])
    return buf.getvalue()


# ──────────────────────────────────────────────────────────────
# Interface
# ──────────────────────────────────────────────────────────────
st.markdown("""
<div class="app-header">
  <h1>🚛 Inventário de Frota</h1>
  <p>Automação de evidências — leitura de placas e chassis via OCR nos PDFs do inventário</p>
</div>
""", unsafe_allow_html=True)

col_pdf, col_xlsx = st.columns(2, gap="large")

with col_pdf:
    st.markdown('<div class="upload-card"><h3>📄 PDFs do Inventário</h3>'
                '<p>Um ou mais arquivos PDF gerados no inventário de frota</p></div>',
                unsafe_allow_html=True)
    pdf_uploads = st.file_uploader(
        "PDFs", type=["pdf"], accept_multiple_files=True,
        label_visibility="collapsed", key="up_pdf",
    )
    if pdf_uploads:
        nomes = ", ".join(f.name for f in pdf_uploads[:3])
        st.caption(f"✅ {len(pdf_uploads)} arquivo(s): {nomes}"
                   + ("…" if len(pdf_uploads) > 3 else ""))

with col_xlsx:
    st.markdown('<div class="upload-card"><h3>📊 Planilha XLSX</h3>'
                '<p>Planilha de inventário sem a coluna Evidência preenchida</p></div>',
                unsafe_allow_html=True)
    xlsx_upload = st.file_uploader(
        "XLSX", type=["xlsx"], accept_multiple_files=False,
        label_visibility="collapsed", key="up_xlsx",
    )
    if xlsx_upload:
        st.caption(f"✅ {xlsx_upload.name}")

st.markdown("<hr>", unsafe_allow_html=True)

pronto = bool(pdf_uploads and xlsx_upload)
btn    = st.button("▶  Processar Inventário", type="primary",
                   use_container_width=True, disabled=not pronto)

if not pronto:
    st.info("⬆️ Envie os PDFs e a planilha XLSX para habilitar o processamento.")

# ──────────────────────────────────────────────────────────────
# Processamento
# ──────────────────────────────────────────────────────────────
if btn and pronto:

    # 1. Carregar planilha
    with st.status("📂 Carregando planilha…", expanded=False) as s:
        xlsx_bytes = xlsx_upload.read()
        wb_tmp = load_workbook(io.BytesIO(xlsx_bytes))
        ws_tmp = wb_tmp["Inventário"]
        placa_lkp, chassis_lkp = construir_lookup(ws_tmp)
        total_reg = sum(1 for r in range(4, ws_tmp.max_row + 1)
                        if ws_tmp.cell(r, 1).value)
        s.update(label=f"✅ {total_reg} registros carregados — "
                       f"{len(placa_lkp)} placas · {len(chassis_lkp)} chassis indexados",
                 state="complete")

    # 2. Extrair imagens
    with st.status("🖼️ Extraindo imagens dos PDFs…", expanded=False) as s:
        imagens = extrair_imagens(pdf_uploads)
        s.update(label=f"✅ {len(imagens)} imagens extraídas de {len(pdf_uploads)} PDF(s)",
                 state="complete")

    if not imagens:
        st.error("❌ Nenhuma imagem encontrada nos PDFs.")
        st.stop()

    # 3. OCR + correspondência
    st.markdown("#### 🔍 Lendo placas e chassis…")
    barra   = st.progress(0.0)
    legenda = st.empty()

    linhas_encontradas = set()
    nao_identificadas  = []
    log_matches        = []

    for i, img in enumerate(imagens):
        legenda.caption(f"Imagem {i+1} / {len(imagens)}  ·  "
                        f"{img['src']}  ·  pág. {img['pag']}")
        try:
            candidatos, concat = ocr_candidatos(img["pil"])
            identificador, linha = buscar_na_planilha(
                candidatos, concat, placa_lkp, chassis_lkp
            )
            if linha:
                linhas_encontradas.add(linha)
                log_matches.append({
                    "Identificador": identificador,
                    "Linha XLSX": linha,
                    "PDF": img["src"],
                    "Página": img["pag"],
                })
            else:
                nao_identificadas.append(img)
        except Exception:
            nao_identificadas.append(img)

        barra.progress((i + 1) / len(imagens))

    barra.empty()
    legenda.empty()

    # 4. Salvar XLSX
    with st.status("💾 Atualizando coluna Evidência…", expanded=False) as s:
        xlsx_saida = atualizar_xlsx(xlsx_bytes, linhas_encontradas)
        s.update(label=f"✅ {len(linhas_encontradas)} linhas marcadas como SIM",
                 state="complete")

    # 5. ZIP não identificadas
    zip_bytes = montar_zip(nao_identificadas) if nao_identificadas else None

    st.session_state.update({
        "xlsx_saida":    xlsx_saida,
        "zip_bytes":     zip_bytes,
        "n_imgs":        len(imagens),
        "n_sim":         len(linhas_encontradas),
        "n_nao":         len(nao_identificadas),
        "log":           log_matches,
    })
    st.success("✅ Processamento concluído!")

# ──────────────────────────────────────────────────────────────
# Resultados
# ──────────────────────────────────────────────────────────────
if "xlsx_saida" in st.session_state:
    st.markdown("<hr>", unsafe_allow_html=True)

    ni, ns, nn = (st.session_state["n_imgs"],
                  st.session_state["n_sim"],
                  st.session_state["n_nao"])

    st.markdown(f"""
    <div class="metrics">
      <div class="metric-box">
        <div class="val">{ni}</div>
        <div class="lbl">Imagens processadas</div>
      </div>
      <div class="metric-box green">
        <div class="val">{ns}</div>
        <div class="lbl">Registros com SIM</div>
      </div>
      <div class="metric-box {'orange' if nn else 'green'}">
        <div class="val">{nn}</div>
        <div class="lbl">Não identificadas</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<hr>", unsafe_allow_html=True)

    c1, c2 = st.columns(2, gap="large")

    with c1:
        st.download_button(
            label="📥  Download Planilha Preenchida (.xlsx)",
            data=st.session_state["xlsx_saida"],
            file_name="INVENTARIO_FROTA_EVIDENCIAS.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )

    with c2:
        if st.session_state["zip_bytes"]:
            st.download_button(
                label="📦  Download Imagens Não Identificadas (.zip)",
                data=st.session_state["zip_bytes"],
                file_name="nao_identificadas.zip",
                mime="application/zip",
                use_container_width=True,
            )
        else:
            st.success("🎉 Todas as imagens foram identificadas!")

    if st.session_state.get("log"):
        with st.expander(f"🔎 Detalhes — {ns} registros identificados"):
            st.dataframe(
                pd.DataFrame(st.session_state["log"]),
                use_container_width=True,
                hide_index=True,
            )
