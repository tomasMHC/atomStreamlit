import streamlit as st
import pandas as pd
import base64
from io import BytesIO
import requests
import unicodedata
from pathlib import Path

st.set_page_config(page_title="PharmaGroup katalóg", layout="wide")

# =========================================================
# Session state
# =========================================================
if "setup_done" not in st.session_state:
    st.session_state.setup_done = False

if "file_bytes" not in st.session_state:
    st.session_state.file_bytes = None

if "data" not in st.session_state:
    st.session_state.data = None

if "cart" not in st.session_state:
    st.session_state.cart = []

if "selected_sheets" not in st.session_state:
    st.session_state.selected_sheets = []

if "item_col" not in st.session_state:
    st.session_state.item_col = None

if "category_col" not in st.session_state:
    st.session_state.category_col = None

if "price_col" not in st.session_state:
    st.session_state.price_col = None

if "desc_col" not in st.session_state:
    st.session_state.desc_col = None

if "autoload_enabled" not in st.session_state:
    st.session_state.autoload_enabled = True

if "next_cart_id" not in st.session_state:
    st.session_state.next_cart_id = 1


# =========================================================
# Styling
# =========================================================
st.markdown("""
<style>
header[data-testid="stHeader"] { display: none; }
[data-testid="stToolbar"] { display: none !important; }
.stAppDeployButton { display: none !important; }
#MainMenu { visibility: hidden; }
footer { visibility: hidden; }
[data-testid="stStatusWidget"] { display: none !important; }
[data-testid="stDecoration"] { display: none !important; }

.block-container {
    overflow: visible !important;
    padding-top: 1.2rem;
    padding-bottom: 1.2rem;
}

/* sticky cart */
.sticky-cart-container {
    position: sticky;
    top: 12px;
    z-index: 50;
}
</style>
""", unsafe_allow_html=True)


# =========================================================
# Helpers
# =========================================================
def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    output.seek(0)
    return output.getvalue()


@st.cache_data(show_spinner=False)
def download_excel_from_private_url():
    excel_url = st.secrets["excel_url"]

    headers = {}
    if "excel_auth_header" in st.secrets and st.secrets["excel_auth_header"]:
        headers["Authorization"] = st.secrets["excel_auth_header"]

    response = requests.get(excel_url, headers=headers, timeout=60)
    response.raise_for_status()

    return response.content


def normalize_text(text):
    text = str(text).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return text


def guess_column(columns, candidates):
    normalized = {normalize_text(c): c for c in columns}
    for cand in candidates:
        cand_norm = normalize_text(cand)
        if cand_norm in normalized:
            return normalized[cand_norm]
    return None


@st.cache_data
def get_excel_sheets(file_bytes):
    xl = pd.ExcelFile(BytesIO(file_bytes), engine="openpyxl")
    return xl.sheet_names


@st.cache_data
def load_selected_sheets(file_bytes, selected_sheets):
    sheets_dict = pd.read_excel(
        BytesIO(file_bytes),
        sheet_name=selected_sheets,
        engine="openpyxl"
    )

    frames = []
    for sheet_name, df in sheets_dict.items():
        if df is None or df.empty:
            continue

        temp = df.copy()
        temp["__sheet__"] = sheet_name
        frames.append(temp)

    if not frames:
        return pd.DataFrame()

    return pd.concat(frames, ignore_index=True)


@st.cache_data
def build_standard_table(df, item_col, category_col, price_col, desc_col=None):
    rename_map = {
        item_col: "item",
        category_col: "category",
        price_col: "price"
    }

    if desc_col:
        rename_map[desc_col] = "description"

    out = df.rename(columns=rename_map).copy()

    required = ["item", "category", "price", "__sheet__"]
    if desc_col:
        required.append("description")

    out = out[required].copy()

    out["item"] = out["item"].astype(str).str.strip()
    out["category"] = out["category"].astype(str).str.strip()
    out["price"] = pd.to_numeric(out["price"], errors="coerce")

    out = out.dropna(subset=["item", "category", "price"])
    out = out[out["item"] != ""]
    out = out[out["category"] != ""]

    out = out.rename(columns={"__sheet__": "eqp_type"})
    out = out.reset_index(drop=True)

    return out


def add_to_cart(item_row, qty):
    item_name = str(item_row["item"])
    category = str(item_row["category"])
    price = float(item_row["price"])
    eqp_type = str(item_row["eqp_type"])

    for cart_item in st.session_state.cart:
        if (
            cart_item["item"] == item_name
            and cart_item["category"] == category
            and cart_item["price"] == price
            and cart_item["eqp_type"] == eqp_type
        ):
            cart_item["qty"] += int(qty)
            return

    st.session_state.cart.append({
        "cart_id": st.session_state.next_cart_id,
        "item": item_name,
        "category": category,
        "price": price,
        "qty": int(qty),
        "eqp_type": eqp_type
    })

    st.session_state.next_cart_id += 1


def get_cart_df():
    if not st.session_state.cart:
        return pd.DataFrame(columns=[
            "item", "category", "eqp_type", "price", "qty", "line_total"
        ])

    df = pd.DataFrame(st.session_state.cart)
    df["line_total"] = df["price"] * df["qty"]

    return df[["item", "category", "eqp_type", "price", "qty", "line_total"]]


def reset_setup(keep_cart=False):
    st.session_state.setup_done = False
    st.session_state.file_bytes = None
    st.session_state.data = None
    st.session_state.selected_sheets = []
    st.session_state.item_col = None
    st.session_state.category_col = None
    st.session_state.price_col = None
    st.session_state.desc_col = None

    if not keep_cart:
        st.session_state.cart = []
        st.session_state.next_cart_id = 1


def load_logo(path):
    p = Path(path)
    if not p.exists():
        return None

    with open(p, "rb") as f:
        return base64.b64encode(f.read()).decode()


def try_autoload_default_excel():
    file_bytes = download_excel_from_private_url()
    sheet_names = get_excel_sheets(file_bytes)

    raw_df = load_selected_sheets(file_bytes, sheet_names)
    if raw_df.empty:
        raise ValueError("V súkromnom Excel súbore sa nenašli žiadne dáta.")

    available_cols = [c for c in raw_df.columns if c != "__sheet__"]

    item_guess = guess_column(
        available_cols,
        ["polozky", "polozka", "item", "name", "product", "part", "material", "description"]
    )
    category_guess = guess_column(
        available_cols,
        ["kategoria", "category", "group", "type", "family", "class"]
    )
    price_guess = guess_column(
        available_cols,
        ["cena", "price", "cost", "unit price (€)", "unit price", "amount", "value"]
    )
    desc_guess = guess_column(
        available_cols,
        ["popis", "description", "detail", "info"]
    )

    if not item_guess or not category_guess or not price_guess:
        raise ValueError(
            f"Nepodarilo sa automaticky nájsť potrebné stĺpce. Nájdené stĺpce: {available_cols}"
        )

    data = build_standard_table(raw_df, item_guess, category_guess, price_guess, desc_guess)
    if data.empty:
        raise ValueError("Po normalizácii sa nenašli žiadne platné riadky.")

    st.session_state.file_bytes = file_bytes
    st.session_state.data = data
    st.session_state.selected_sheets = sheet_names
    st.session_state.item_col = item_guess
    st.session_state.category_col = category_guess
    st.session_state.price_col = price_guess
    st.session_state.desc_col = desc_guess
    st.session_state.setup_done = True


# =========================================================
# Title
# =========================================================
logo_base64 = load_logo("main_logo.png")

if logo_base64:
    st.markdown(
        f"""
        <div style="display:flex;align-items:center;gap:15px;">
            <img src="data:image/png;base64,{logo_base64}" width="100">
            <div>
                <h1 style="margin-bottom:0;">PharmaGroup katalóg</h1>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )
else:
    st.title("PharmaGroup katalóg")


# =========================================================
# Setup (autoload + manuálne)
# =========================================================
if not st.session_state.setup_done and st.session_state.autoload_enabled:
    try:
        try_autoload_default_excel()
    except Exception as e:
        st.warning(f"Automatické načítanie súkromného Excel súboru zlyhalo: {e}")

if not st.session_state.setup_done:
    st.subheader("Nastavenie")

    uploaded_file = st.file_uploader(
        "Nahraj Excel súbor",
        type=["xlsx", "xls"],
        help="Nahraj Excel súbor s jedným alebo viacerými hárkami."
    )

    if uploaded_file is None:
        st.info("Najprv nahraj Excel súbor.")
        st.stop()

    file_bytes = uploaded_file.getvalue()

    try:
        sheet_names = get_excel_sheets(file_bytes)
    except Exception as e:
        st.error(f"Nepodarilo sa načítať Excel súbor: {e}")
        st.stop()

    selected_sheets = st.multiselect(
        "Vyber hárky",
        options=sheet_names,
        default=sheet_names
    )

    if not selected_sheets:
        st.warning("Vyber aspoň jeden hárok.")
        st.stop()

    try:
        raw_df = load_selected_sheets(file_bytes, selected_sheets)
    except Exception as e:
        st.error(f"Chyba pri načítaní hárkov: {e}")
        st.stop()

    if raw_df.empty:
        st.warning("V zvolených hárkoch sa nenašli žiadne dáta.")
        st.stop()

    available_cols = [c for c in raw_df.columns if c != "__sheet__"]

    if not available_cols:
        st.error("V zvolených hárkoch sa nenašli použiteľné stĺpce.")
        st.stop()

    st.markdown("### Namapuj stĺpce")

    item_guess = guess_column(
        available_cols,
        ["polozka", "polozky", "item", "name", "product", "part", "part name", "material", "description"]
    )
    category_guess = guess_column(
        available_cols,
        ["kategoria", "category", "group", "type", "family", "class"]
    )
    price_guess = guess_column(
        available_cols,
        ["cena", "price", "cost", "unit price", "unit price (€)", "amount", "value"]
    )
    desc_guess = guess_column(
        available_cols,
        ["popis", "description", "detail", "info"]
    )

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        item_col = st.selectbox(
            "Stĺpec položky",
            options=available_cols,
            index=available_cols.index(item_guess) if item_guess in available_cols else 0,
            help="Stĺpec s názvom položky / produktu / dielu."
        )

    with c2:
        category_col = st.selectbox(
            "Stĺpec kategórie",
            options=available_cols,
            index=available_cols.index(category_guess) if category_guess in available_cols else 0,
            help="Stĺpec s kategóriou / skupinou / typom."
        )

    with c3:
        price_col = st.selectbox(
            "Stĺpec ceny",
            options=available_cols,
            index=available_cols.index(price_guess) if price_guess in available_cols else 0,
            help="Stĺpec s číselnou cenou položky."
        )

    with c4:
        desc_col = st.selectbox(
            "Stĺpec popisu (voliteľné)",
            options=["(žiadny)"] + available_cols,
            index=(available_cols.index(desc_guess) + 1) if desc_guess in available_cols else 0,
            help="Stĺpec s textovým popisom položky."
        )

    with st.expander("Náhľad nahraných dát"):
        st.dataframe(raw_df, use_container_width=True)

    if st.button("Potvrdiť nastavenie", type="primary"):
        try:
            desc_final = None if desc_col == "(žiadny)" else desc_col
            data = build_standard_table(raw_df, item_col, category_col, price_col, desc_final)
        except Exception as e:
            st.error(f"Mapovanie stĺpcov zlyhalo: {e}")
            st.stop()

        if data.empty:
            st.warning("Po aplikovaní zvolených stĺpcov sa nenašli žiadne platné riadky.")
            st.stop()

        st.session_state.file_bytes = file_bytes
        st.session_state.data = data
        st.session_state.selected_sheets = selected_sheets
        st.session_state.item_col = item_col
        st.session_state.category_col = category_col
        st.session_state.price_col = price_col
        st.session_state.desc_col = desc_final
        st.session_state.setup_done = True

    st.stop()


# =========================================================
# Po nastavení – hlavná aplikácia
# =========================================================
data = st.session_state.data

if data is None or data.empty:
    st.error("Dáta nie sú dostupné. Resetuj a nahraj súbor znova.")
    st.stop()

# ---------------------------------------------------------
# Filtre
# ---------------------------------------------------------
filter1, filter2, filter3 = st.columns([2, 2, 3])

filtered = data.copy()

with filter1:
    eqp_options = ["Všetko"] + sorted(data["eqp_type"].dropna().unique().tolist())
    selected_eqp = st.selectbox("Typ vybavenia", eqp_options)

if selected_eqp != "Všetko":
    filtered = filtered[filtered["eqp_type"] == selected_eqp].copy()

with filter2:
    category_options = ["Všetko"] + sorted(filtered["category"].dropna().unique().tolist())
    selected_category = st.selectbox("Kategória", category_options)

if selected_category != "Všetko":
    filtered = filtered[filtered["category"] == selected_category].copy()

with filter3:
    search_text = st.text_input("Hľadať položku")

if search_text:
    filtered = filtered[
        filtered["item"].str.contains(search_text, case=False, na=False)
    ].copy()

filtered = filtered.sort_values(["category", "item"]).reset_index(drop=True)


# =========================================================
# Layout – položky + košík
# =========================================================
left_col, right_col = st.columns([2.4, 1.6], gap="large")

# ---------------------------------------------------------
# Položky – rýchle renderovanie cez dataframe
# ---------------------------------------------------------
with left_col:
    st.markdown("## Dostupné položky")

    if filtered.empty:
        st.warning("Žiadne položky nevyhovujú aktuálnym filtrom.")
    else:
        display_df = filtered[["item", "eqp_type", "category", "price"]].copy()
        display_df["price"] = display_df["price"].round(2)
        display_df = display_df.reset_index().rename(columns={"index": "ID"})

        st.dataframe(display_df, use_container_width=True)

        selected_id = st.number_input(
            "ID položky na pridanie",
            min_value=0,
            max_value=len(display_df) - 1,
            step=1
        )

        add_qty = st.number_input(
            "Množstvo",
            min_value=1,
            max_value=1000,
            step=1,
            value=1
        )

        if st.button("Pridať do košíka", use_container_width=True):
            add_to_cart(filtered.iloc[selected_id], add_qty)
            st.toast("Položka pridaná do košíka")


# ---------------------------------------------------------
# Košík – data_editor + export
# ---------------------------------------------------------
with right_col:
    st.markdown('<div class="sticky-cart-container">', unsafe_allow_html=True)

    st.markdown("""
    <div class="cart-header">
        <div class="cart-header-title">🛒 Vybraté položky</div>
        <div class="cart-header-subtitle">Prehľad položiek na stiahnutie / objednanie</div>
    </div>
    """, unsafe_allow_html=True)

    cdf = get_cart_df()

    if cdf.empty:
        st.info("Zoznam položiek je prázdny.")
    else:
        # Košík ako editable tabuľka
        edit_df = cdf.copy()
        edit_df = edit_df.rename(columns={
            "item": "Položka",
            "category": "Kategória",
            "eqp_type": "Typ vybavenia",
            "price": "Cena bez DPH (€)",
            "qty": "Množstvo",
            "line_total": "Celkom bez DPH (€)"
        })

        edited = st.data_editor(
            edit_df,
            num_rows="dynamic",
            use_container_width=True,
            key="cart_editor"
        )

        # Aktualizácia session_state.cart podľa edited
        new_cart = []
        for _, row in edited.iterrows():
            if pd.isna(row["Položka"]) or pd.isna(row["Kategória"]):
                continue
            new_cart.append({
                "cart_id": None,
                "item": row["Položka"],
                "category": row["Kategória"],
                "eqp_type": row["Typ vybavenia"],
                "price": float(row["Cena bez DPH (€)"]) if not pd.isna(row["Cena bez DPH (€)"]) else 0.0,
                "qty": int(row["Množstvo"]) if not pd.isna(row["Množstvo"]) else 0
            })

        st.session_state.cart = new_cart
        cdf = get_cart_df()

        total = cdf["line_total"].sum()
        total_w_dph = total * 1.23

        with st.container(border=True):
            r1c1, r1c2 = st.columns([2.2, 1])
            with r1c1:
                st.markdown("Celkom bez DPH")
            with r1c2:
                st.markdown(
                    f"<div style='text-align:right; font-weight:600;'>{total:.2f} €</div>",
                    unsafe_allow_html=True
                )

            st.divider()

            r2c1, r2c2 = st.columns([2.2, 1])
            with r2c1:
                st.markdown(
                    "<div style='font-size:18px; font-weight:700;'>Celkom s DPH</div>",
                    unsafe_allow_html=True
                )
            with r2c2:
                st.markdown(
                    f"<div style='text-align:right; font-size:26px; font-weight:800; color:#0f766e; white-space:nowrap;'>{total_w_dph:.2f} €</div>",
                    unsafe_allow_html=True
                )

        btn1, btn2 = st.columns(2)
        with btn1:
            if st.button("Vymazať zoznam", use_container_width=True):
                st.session_state.cart = []
                st.session_state.next_cart_id = 1
                st.toast("Košík vymazaný")

        with btn2:
            export_df = cdf.copy()
            export_df = export_df.rename(columns={
                "item": "Položka",
                "category": "Kategória",
                "eqp_type": "Typ vybavenia",
                "price": "Cena bez DPH (€)",
                "qty": "Množstvo",
                "line_total": "Celkom bez DPH (€)"
            })

            total_rows = pd.DataFrame([
                {
                    "Položka": "",
                    "Kategória": "",
                    "Typ vybavenia": "",
                    "Cena bez DPH (€)": "",
                    "Množstvo": "Celkom bez DPH",
                    "Celkom": round(total, 2)
                },
                {
                    "Položka": "",
                    "Kategória": "",
                    "Typ vybavenia": "",
                    "Cena bez DPH (€)": "",
                    "Množstvo": "Celkom s DPH",
                    "Celkom": round(total_w_dph, 2)
                }
            ])

            export_df = pd.concat([export_df, total_rows], ignore_index=True)

            excel_data = to_excel_bytes(export_df)

            st.download_button(
                label="Stiahnuť Excel",
                data=excel_data,
                file_name="zoznam.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    st.markdown('</div>', unsafe_allow_html=True)