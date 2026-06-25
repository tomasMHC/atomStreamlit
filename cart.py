import streamlit as st
import pandas as pd
import base64
from io import BytesIO
import requests
import unicodedata
from pathlib import Path
MAX_ITEMS = 50
st.set_page_config(page_title="PharmaGroup katalóg", layout="wide")

# =========================================================
# Session state initialization
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

div[data-baseweb="input"] { min-width: 80px !important; }

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

/* LEFT side section */
.items-panel {
    padding-right: 10px;
}

/* kompaktnejšie riadky v zozname položiek */
.items-panel .stColumn {
    padding-top: 0px !important;
    padding-bottom: 0px !important;
}

/* kompaktnejší popis položky */
.item-description {
    font-size: 13px;
    color: gray;
    margin-top: -2px;
    margin-bottom: 2px;
}

/* RIGHT side cart card */
.cart-panel {
    background: #f8fafc;
    border: 1px solid #e5e7eb;
    border-radius: 18px;
    padding: 18px 18px 16px 18px;
    box-shadow: 0 4px 14px rgba(15, 23, 42, 0.06);
}

/* cart header strip */
.cart-header {
    background: linear-gradient(90deg, #eff6ff 0%, #f8fafc 100%);
    border: 1px solid #dbeafe;
    border-radius: 12px;
    padding: 10px 14px;
    margin-bottom: 12px;
}

.cart-header-title {
    font-size: 20px;
    font-weight: 700;
    color: #0f172a;
    margin: 0;
}

.cart-header-subtitle {
    font-size: 12px;
    color: #64748b;
    margin-top: 2px;
}

/* subtle separator feel between columns on desktop */
@media (min-width: 900px) {
    .cart-panel {
        margin-left: 10px;
    }
}

/* mobile */
@media (max-width: 768px) {
    .cart-panel {
        margin-top: 14px;
        padding: 14px;
        border-radius: 14px;
    }

    .cart-header-title {
        font-size: 18px;
    }
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

@st.cache_data(show_spinner=False)
def get_excel_sheets(file_bytes):
    xl = pd.ExcelFile(BytesIO(file_bytes), engine="openpyxl")
    return xl.sheet_names

@st.cache_data(show_spinner=False)
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

        df["__sheet__"] = sheet_name   # netreba temp = df.copy()
        frames.append(df)

    if not frames:
        return pd.DataFrame()

    return pd.concat(frames, ignore_index=True)

@st.cache_data(show_spinner=False)
def build_standard_table(df, item_col, category_col, price_col, desc_col=None):
    rename_map = {
        item_col: "item",
        category_col: "category",
        price_col: "price"
    }

    if desc_col:
        rename_map[desc_col] = "description"

    required = ["item", "category", "price", "__sheet__"]
    if desc_col:
        required.append("description")

    out = df.rename(columns=rename_map)[required]

    out["item"] = out["item"].astype(str).str.strip()
    out["category"] = out["category"].astype(str).str.strip()
    out["price"] = pd.to_numeric(out["price"], errors="coerce")

    out = out.dropna(subset=["item", "category", "price"])
    out = out[(out["item"] != "") & (out["category"] != "")]
    out = out.rename(columns={"__sheet__": "eqp_type"}).reset_index(drop=True)

    return out

@st.cache_data(show_spinner=False)
def load_logo(path):
    p = Path(path)
    if not p.exists():
        return None

    with open(p, "rb") as f:
        return base64.b64encode(f.read()).decode()

@st.cache_data(show_spinner=False)
def build_export_excel(cart_records, total, total_w_dph):
    cdf = pd.DataFrame(cart_records)
    if not cdf.empty:
        cdf["line_total"] = cdf["price"] * cdf["qty"]
        cdf["line_total"] = cdf["line_total"].round(2)
        cdf["price"] = cdf["price"].round(2)

    export_df = cdf.copy()
    export_df = export_df.rename(columns={
        "item": "Položka",
        "category": "Kategória",
        "eqp_type": "Typ vybavenia",
        "price": "Cena bez DPH (€)",
        "qty": "Množstvo",
        "line_total": "Celkom"
    })

    export_df = export_df[["Položka", "Kategória", "Typ vybavenia", "Cena bez DPH (€)", "Množstvo", "Celkom"]]

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
    return to_excel_bytes(export_df)


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


def get_excel_sheets(file_bytes):
    xl = pd.ExcelFile(BytesIO(file_bytes), engine="openpyxl")
    return xl.sheet_names


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

def make_cart_key(item_name, category, price, eqp_type):
    return (item_name, category, float(price), eqp_type)


if "cart_lookup" not in st.session_state:
    st.session_state.cart_lookup = {}


def add_to_cart(item_row, qty):
    item_name = str(item_row["item"])
    category = str(item_row["category"])
    price = float(item_row["price"])
    eqp_type = str(item_row["eqp_type"])
    qty = int(qty)

    key = make_cart_key(item_name, category, price, eqp_type)

    if key in st.session_state.cart_lookup:
        cart_id = st.session_state.cart_lookup[key]
        for item in st.session_state.cart:
            if item["cart_id"] == cart_id:
                item["qty"] += qty
                return

    cart_id = st.session_state.next_cart_id
    st.session_state.cart.append({
        "cart_id": cart_id,
        "item": item_name,
        "category": category,
        "price": price,
        "qty": qty,
        "eqp_type": eqp_type
    })
    st.session_state.cart_lookup[key] = cart_id
    st.session_state.next_cart_id += 1


def get_cart_df():
    if not st.session_state.cart:
        return pd.DataFrame(columns=[
            "cart_id", "item", "category", "eqp_type", "price", "qty", "line_total"
        ])

    df = pd.DataFrame(st.session_state.cart)
    df["line_total"] = df["price"] * df["qty"]

    return df[["cart_id", "item", "category", "eqp_type", "price", "qty", "line_total"]]


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
        clear_cart_qty_widget_state()


def update_cart_qty(cart_id, new_qty):
    new_qty = int(new_qty)

    for idx, cart_item in enumerate(st.session_state.cart):
        if cart_item["cart_id"] == cart_id:
            if new_qty < 1:
                st.session_state.cart.pop(idx)
                clear_cart_qty_widget_state()
            else:
                st.session_state.cart[idx]["qty"] = new_qty
            return

def remove_from_cart(cart_id):
    for idx, cart_item in enumerate(st.session_state.cart):
        if cart_item["cart_id"] == cart_id:
            st.session_state.cart.pop(idx)
            qty_key = f"cart_qty_{cart_id}"
            if qty_key in st.session_state:
                del st.session_state[qty_key]
            return
        
def clear_cart_qty_widget_state():
    for key in list(st.session_state.keys()):
        if str(key).startswith("cart_qty_"):
            del st.session_state[key]

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
    # st.caption("Nahraj Excel súbor, namapuj stĺpce a používaj filtre a košík.")


# =========================================================
# Setup section (auto-load + manuálne nastavenie)
# =========================================================
if not st.session_state.setup_done and st.session_state.autoload_enabled:
    try:
        try_autoload_default_excel()
        st.rerun()
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

        st.rerun()

    st.stop()


# =========================================================
# Compact mode (after setup)
# =========================================================
data = st.session_state.data

if data is None or data.empty:
    st.error("Dáta nie sú dostupné. Resetuj a nahraj súbor znova.")
#     if st.button("Resetovať"):
#         reset_setup(keep_cart=False)
#         st.rerun()
#     st.stop()


# # Top compact controls
# top_left, top_mid, top_right = st.columns([3, 2, 1])

# with top_left:
#     st.success("Excel načítaný")

# with top_mid:
#     file_info = f"Hárky: {', '.join(st.session_state.selected_sheets)}"
#     st.caption(file_info)

# with top_right:
#     if st.button("Zmeniť súbor / mapovanie"):
#         reset_setup(keep_cart=False)
#         st.session_state.autoload_enabled = False
#         st.rerun()


# =========================================================
# Compact dropdown filters
# =========================================================
filter1, filter2, filter3 = st.columns([2, 2, 3])

filtered = data

with filter1:
    eqp_options = ["Všetko"] + sorted(data["eqp_type"].dropna().unique().tolist())
    selected_eqp = st.selectbox("Typ vybavenia", eqp_options)

if selected_eqp != "Všetko":
    filtered = filtered[filtered["eqp_type"] == selected_eqp]

with filter2:
    category_options = ["Všetko"] + sorted(filtered["category"].dropna().unique().tolist())
    selected_category = st.selectbox("Kategória", category_options)

if selected_category != "Všetko":
    filtered = filtered[filtered["category"] == selected_category]

with filter3:
    search_text = st.text_input("Hľadať položku")

if search_text:
    filtered = filtered[filtered["item"].str.contains(search_text, case=False, na=False)]

filtered = filtered.sort_values(["category", "item"]).reset_index(drop=True)


# =========================================================
# Main layout
# =========================================================
left_col, right_col = st.columns([2.4, 1.6], gap = "large")

# ---------------------------------------------------------
# Available items
# ---------------------------------------------------------
with left_col:
    st.markdown('<div class="items-panel">', unsafe_allow_html=True)
    st.markdown("## Dostupné položky")

    if filtered.empty:
        st.warning("Žiadne položky nevyhovujú aktuálnym filtrom.")
    else:
        h1, h2, h3, h4, h5 = st.columns([4, 2, 2, 1.3, 1.2])
        with h1:
            st.markdown("**Položka**")
        with h2:
            st.markdown("**Typ vybavenia**")
        with h3:
            st.markdown("**Kategória**")
        with h4:
            st.markdown("**Cena bez DPH (€)**")

        st.divider()

        for row in filtered.itertuples(index=True):
            c1, c2, c3, c4, c5 = st.columns([4, 2, 2, 1.3, 1.2])

            with c1:
                st.markdown(f"**{row.item}**")

            with c2:
                st.write(row.eqp_type)

            with c3:
                st.write(row.category)

            with c4:
                st.write(f"{row.price:.2f}")

            with c5:
                if st.button("Pridať", key=f"add_{row.Index}", use_container_width=True):
                    add_to_cart({
                        "item": row.item,
                        "eqp_type": row.eqp_type,
                        "category": row.category,
                        "price": row.price
                    },1)
                    st.toast("Položka pridaná")

            if hasattr(row, "description") and pd.notna(row.description) and str(row.description).strip() != "":
                d1, _ = st.columns([6, 5])
                with d1:
                    st.markdown(
                        f"""
                        <div class="item-description">
                            {row.description}
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

            st.markdown("<hr style='margin:6px 0;'>", unsafe_allow_html=True)



# ---------------------------------------------------------
# Cart (sticky in right column)
# ---------------------------------------------------------

def qty_changed(cart_id):
    key = f"cart_qty_{cart_id}"
    if key not in st.session_state:
        return
    new_qty = st.session_state[key]
    update_cart_qty(cart_id, new_qty)

with right_col:
    st.markdown('<div class="cart-panel">', unsafe_allow_html=True)
    
    with st.container(border=True):
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
            h1, h2, h3, h4, h5, h6 = st.columns([2.4, 1.2, 1.3, 1.0, 1.3, 1.2])
            with h1:
                st.markdown("**Položka**")
            with h2:
                st.markdown("**Typ**")
            with h3:
                st.markdown("**Cena (€)**")
            with h4:
                st.markdown("**Množ.**")
            with h5:
                st.markdown("**Spolu (€)**")
            with h6:
                st.markdown("")

            st.divider()



            for _, row in cdf.iterrows():
                a, b, c, d, e, f = st.columns([2.4, 1.2, 1.3, 1.0, 1.3, 1.2])
                cart_id = int(row["cart_id"])

                with a:
                    st.markdown(
                        f"""
                        <div style="line-height:1.1;">
                            <div style="font-weight:600;">{row['item']}</div>
                            <div style="font-size:12px; color:gray;">{row['category']}</div>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

                with b:
                    st.write(row["eqp_type"])

                with c:
                    st.write(f"{row['price']:.2f} €")

                with d:
                    qty_key = f"cart_qty_{cart_id}"
                    if qty_key not in st.session_state:
                        st.session_state[qty_key] = int(row["qty"])

                    st.number_input(
                        "Množstvo",
                        min_value=1,
                        max_value=1000,
                        step=1,
                        key=qty_key,
                        label_visibility="collapsed",
                        on_change=qty_changed,
                        args=(cart_id,)
                    )

                with e:
                    st.markdown(
                        f"""
                        <div style="text-align:right; font-weight:600; padding-top:6px;">
                            {row['line_total']:.2f} €
                        </div>
                        <div style="text-align:right; font-size:11px; color:gray;">
                            spolu
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

                with f:
                    if st.button("Odstrániť", key=f"remove_{cart_id}", use_container_width=True):
                        remove_from_cart(cart_id)
                        st.toast("Položka odstránená")
                        st.rerun()

                st.divider()

            # po prípadných zmenách v košíku prepočítaj cdf
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
                    clear_cart_qty_widget_state()
                    st.toast("Zoznam vymazaný")
                    
            with btn2:
                # Export do Excelu – s riadkami CELKOM
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
                export_df = export_df[["Položka", "Kategória", "Typ vybavenia", "Cena bez DPH (€)", "Množstvo", "Celkom"]]
                excel_data = build_export_excel(st.session_state.cart, total, total_w_dph)


                st.download_button(
                    label="Stiahnuť Excel",
                    data=excel_data,
                    file_name="zoznam.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

    st.markdown('</div>', unsafe_allow_html=True)




# =========================================================
# Optional preview
# =========================================================
# with st.expander("Náhľad normalizovaných dát používaných aplikáciou"):
#     st.dataframe(data, use_container_width=True)
