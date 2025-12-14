import io
import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont
import qrcode
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

st.set_page_config(page_title="Пропуска (веб-тест)", layout="wide")
st.title("Пропуска ТСЖ — тест в браузере")

# Координаты под ваш шаблон 1088×768 (можно править)
LAYOUT = {
    "size": (1088, 768),
    "fields": {
        "PassNo":   {"x": 520, "y": 35,  "size": 86},
        "Kpp":      {"x": 705, "y": 150, "size": 76},
        "PlotNo":   {"x": 300, "y": 295, "size": 76},
        "Brand":    {"x": 320, "y": 445, "size": 72},
        "Plate":    {"x": 690, "y": 445, "size": 72},
        "FullName": {"x": 250, "y": 535, "size": 56},
        "EndDate":  {"x": 440, "y": 665, "size": 68},
        "AutoNo":   {"x": 30,  "y": 705, "size": 22}
    },
    "qr": {"x": 900, "y": 585, "size": 160}
}

def kpp_to_str(v):
    if pd.isna(v):
        return ""
    s = str(v).strip()
    return s.replace(".", ",")  # 4.6 -> 4,6

def normalize_plate(v):
    if pd.isna(v):
        return ""
    s = str(v).replace("\u00A0", " ").strip()
    return s.replace(" ", "")

def date_to_str(v):
    if pd.isna(v):
        return ""
    try:
        dt = pd.to_datetime(v)
        return dt.strftime("%d.%m.%Y")
    except:
        return str(v).strip()

def build_qr_payload(row: dict):
    # ВЕСЬ набор полей (ключ=значение;)
    def clean(x):
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return ""
        return str(x).replace("\r", " ").replace("\n", " ").replace("\t", " ").strip()

    parts = [
        ("pass_no",   clean(row.get("№ пропуска", ""))),
        ("plot_no",   clean(row.get("№ участка", ""))),
        ("type",      clean(row.get("Тип пропуска", ""))),
        ("fio",       clean(row.get("ФИО", ""))),
        ("passport",  clean(row.get("Паспорт", ""))),
        ("position",  clean(row.get("Должность", ""))),
        ("brand",     clean(row.get("Марка", ""))),
        ("plate",     clean(normalize_plate(row.get("Гос. Номер", "")))),
        ("kpp",       clean(kpp_to_str(row.get("КПП", "")))),
        ("req_date",  clean(row.get("Дата заявки", ""))),
        ("applicant", clean(row.get("Заявитель", ""))),
        ("end_date",  clean(row.get("Дата окончания", ""))),
        ("patent",    clean(row.get("Патент", ""))),
    ]
    return "".join([f"{k}={v};" for k, v in parts])

def make_qr_image(payload: str, size: int):
    qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_Q, box_size=10, border=2)
    qr.add_data(payload)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
    return img.resize((size, size))

def draw_card(template_img: Image.Image, row: dict, auto_no: int):
    img = template_img.copy().convert("RGB")
    draw = ImageDraw.Draw(img)

    def font(sz):
        try:
            return ImageFont.truetype("arial.ttf", sz)
        except:
            return ImageFont.load_default()

    pass_type = str(row.get("Тип пропуска", "")).strip().lower()

    values = {
        "PassNo": str(row.get("№ пропуска", "")).strip(),
        "Kpp": kpp_to_str(row.get("КПП", "")),
        "PlotNo": str(row.get("№ участка", "")).strip(),
        "Brand": str(row.get("Марка", "")).strip(),
        "Plate": normalize_plate(row.get("Гос. Номер", "")),
        "FullName": str(row.get("ФИО", "")).strip(),
        "EndDate": date_to_str(row.get("Дата окончания", "")),
        "AutoNo": str(auto_no)
    }

    # Правило: Пеший — скрыть марку/номер
    if pass_type == "пеший":
        values["Brand"] = ""
        values["Plate"] = ""

    # Текст
    for key, cfg in LAYOUT["fields"].items():
        text = values.get(key, "")
        if not text:
            continue
        draw.text((cfg["x"], cfg["y"]), text, fill=(0, 0, 0), font=font(cfg["size"]))

    # QR
    payload = build_qr_payload(row)
    qr_cfg = LAYOUT["qr"]
    qr_img = make_qr_image(payload, qr_cfg["size"])
    img.paste(qr_img, (qr_cfg["x"], qr_cfg["y"]))

    return img

def images_to_pdf(images, out_bytesio):
    w, h = images[0].size
    c = canvas.Canvas(out_bytesio, pagesize=(w, h))
    for im in images:
        bio = io.BytesIO()
        im.save(bio, format="PNG")
        bio.seek(0)
        c.drawImage(ImageReader(bio), 0, 0, width=w, height=h)
        c.showPage()
    c.save()

# --- UI ---
col1, col2 = st.columns(2)
with col1:
    excel = st.file_uploader("Excel (.xlsx)", type=["xlsx"])
with col2:
    tpl = st.file_uploader("Шаблон JPG (1088×768)", type=["jpg", "jpeg"])

if not excel or not tpl:
    st.info("Загрузите Excel и JPG-шаблон. Потом появится таблица и экспорт PDF.")
    st.stop()

df = pd.read_excel(excel, sheet_name="Лист2")
df = df.dropna(how="all")  # убрать полностью пустые строки

st.subheader("Таблица (выберите строки галочками)")
search = st.text_input("Поиск (ФИО / госномер / участок / № пропуска)")

df_view = df.copy()
if search.strip():
    s = search.strip().lower()
    mask = df_view.astype(str).apply(lambda row: " | ".join(row.values.astype(str)).lower().find(s) >= 0, axis=1)
    df_view = df_view[mask]

df_view.insert(0, "Печать", True)

edited = st.data_editor(
    df_view,
    use_container_width=True,
    hide_index=True,
    num_rows="fixed"
)

template_img = Image.open(tpl)

st.subheader("Предпросмотр")
max_idx = max(0, len(edited) - 1)
idx = st.number_input("Номер строки (0..)", min_value=0, max_value=max_idx, value=0)
row = edited.iloc[int(idx)].to_dict()
preview = draw_card(template_img, row, auto_no=1)
st.image(preview, use_column_width=True)

st.subheader("Экспорт")
if st.button("Сгенерировать PDF по отмеченным"):
    selected = edited[edited["Печать"] == True]
    if len(selected) == 0:
        st.warning("Ничего не отмечено для печати.")
        st.stop()

    images = []
    for i, (_, r) in enumerate(selected.iterrows(), start=1):
        images.append(draw_card(template_img, r.to_dict(), auto_no=i))

    out = io.BytesIO()
    images_to_pdf(images, out)
    out.seek(0)

    st.download_button(
        "Скачать PDF",
        data=out,
        file_name="passes.pdf",
        mime="application/pdf"
    )
