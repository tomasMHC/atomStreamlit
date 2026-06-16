import streamlit as st
PharmaGroup catalogue", layout="wide")import pandas as pd

# =========================================================
# Session state initialization
# =========================================================
DEFAULT_STATE = {
    "setup_done": False,
    "autoload_enabled": True,
    "file_bytes": None,
    "data": None,
    "cart": [],
    "selected_sheets": [],
    "item_col": None,
    "category_col": None,
    "price_col": None,
}

for k, v in DEFAULT_STATE.items():
    if k not in st.session_state:
        st.session_state[k] = v

# =========================================================
# CSS
# =========================================================
st.markdown("""
<style>
/* --- General spacing --- */
.block-container {
    padding-top: 0.8rem;
    padding-bottom: 1rem;
    max-width: 1200px;
}

/* --- Try to reduce Streamlit chrome --- */
header[data-testid="stHeader"] {
    display: none;
}
[data-testid="stToolbar"] {
    display: none !important;
}
.stAppDeployButton {
    display: none !important;
}
#MainMenu {
    visibility: hidden;
}
footer {
    visibility: hidden;
}
[data-testid="stStatusWidget"] {
    display: none !important;
}
[data-testid="stDecoration"] {
    display: none !important;
}

/* --- Compact widgets --- */
div.stButton > button {
    min-height: 2.2rem;
    padding-top: 0.2rem;
    padding-bottom: 0.2rem;
    border-radius: 0.6rem;
}
div[data-baseweb="input"] input {
    padding-top: 0.3rem !important;
    padding-bottom: 0.3rem !important;
}

/* --- Card text spacing --- */
.pg-card-title {
    font-weight: 600;
    font-size: 1.05rem;
    margin-bottom: 0.2rem;
}
.pg-muted {
    color: #6b7280;
    font-size: 0.9rem;
    line-height: 1.2;
}
.pg-header-wrap {
    display: flex;
    align-items: center;
    gap: 14px;
    margin-bottom: 0.5rem;
}
.pg-header-title {
    margin: 0;
    line-height: 1.1;
}
.pg-header-subtitle {
    margin: 0;
    color: #6b7280;
}
</style>
""", unsafe_allow_html=True)

# =========================================================
# Helpers
# =========================================================
def normalize_text(text: str) -> str:
    """
    Lowercase, trim, and remove accents/diacritics.
    Example: Kategória -> kategoria
    """
    text = str(text).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return text


def guess_column(columns, candidates):
    """
    Try to match a column name using normalized comparison.
    """
    normalized_cols = {normalize_text(c): c for c in columns}
    for cand in candidates:
        cand_norm = normalize_text(cand)
        if cand_norm in normalized_cols:
            return normalized_cols[cand_norm]
    return None


def load_logo_base64(path: str):
    """
    Return base64-encoded image if it exists, otherwise None.
    """
    p = Path(path)
    if not p.exists():
        return None
    with open(p, "rb") as f:
        return base64.b64encode(f.read()).decode()


def get_excel_url_from_secrets():
    """
    Supports both:
      st.secrets["excel_url"]
    and:
      st.secrets["sharepoint"]["excel_url"]
    """
    if "excel_url" in st.secrets:
        return st.secrets["excel_url"]

    if "sharepoint" in st.secrets and "excel_url" in st.secrets["sharepoint"]:
        return st.secrets["sharepoint"]["excel_url"]

    raise KeyError("Could not find Excel URL in Streamlit secrets.")


@st.cache_data(show_spinner=False)
def download_excel_from_private_url():
    """
    Download Excel bytes from a private URL stored in Streamlit secrets.
    """
    excel_url = get_excel_url_from_secrets()

    headers = {}
    if "excel_auth_header" in st.secrets and st.secrets["excel_auth_header"]:
        headers["Authorization"] = st.secrets["excel_auth_header"]
    elif "sharepoint" in st.secrets and "excel_auth_header" in st.secrets["sharepoint"]:
        auth_header = st.secrets["sharepoint"]["excel_auth_header"]
        if auth_header:
            headers["Authorization"] = auth_header

    response = requests.get(excel_url, headers=headers, timeout=60)
    response.raise_for_status()

    content_type = response.headers.get("content-type", "")
    if "html" in content_type.lower():
        raise ValueError(
            f"URL returned HTML instead of Excel (content-type: {content_type}). "
            f"Your SharePoint link probably requires login or is a preview page, not a direct file link."
        )

    return response.content


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


def build_standard_table(df, item_col, category_col, price_col):
    """
    Normalize selected columns into:
      item, category, price, eqp_type
    where eqp_type comes from sheet name.
    """
    out = df.rename(columns={
        item_col: "item",
        category_col: "category",
        price_col: "price"
    }).copy()

    required_cols = ["item", "category", "price", "__sheet__"]
    out = out[required_cols].copy()

    out["item"] = out["item"].astype(str).str.strip()
    out["category"] = out["category"].astype(str).str.strip()
    out["price"] = pd.to_numeric(out["price"], errors="coerce")

    out = out.dropna(subset=["item", "category", "price"])
    out = out[out["item"] != ""]
    out = out[out["category"] != ""]

    out = out.rename(columns={"__sheet__": "eqp_type"})
    out = out.reset_index(drop=True)

    return out


def try_autoload_default_excel():
    """
    Auto-load the private Excel file from SharePoint and build normalized data.
    """
    file_bytes = download_excel_from_private_url()
    sheet_names = get_excel_sheets(file_bytes)

    raw_df = load_selected_sheets(file_bytes, sheet_names)
    if raw_df.empty:
        raise ValueError("No data found in the private Excel file.")

    available_cols = [c for c in raw_df.columns if c != "__sheet__"]

    item_guess = guess_column(
        available_cols,
        [
            "polozky", "položky",
            "item", "name", "product", "part", "part name", "material", "description"
        ]
    )
    category_guess = guess_column(
        available_cols,
        [
            "kategoria", "kategória",
            "category", "group", "type", "family", "class"
        ]
    )
    price_guess = guess_column(
        available_cols,
        [
            "cena",
            "price", "cost", "unit price", "amount", "value"
        ]
    )

    if not item_guess or not category_guess or not price_guess:
        raise ValueError(
            f"Could not auto-detect required columns. Found columns: {available_cols}"
        )

    data = build_standard_table(raw_df, item_guess, category_guess, price_guess)
    if data.empty:
        raise ValueError("No valid rows found after normalization.")

    st.session_state.file_bytes = file_bytes
    st.session_state.data = data
    st.session_state.selected_sheets = sheet_names
    st.session_state.item_col = item_guess
    st.session_state.category_col = category_guess
    st.session_state.price_col = price_guess
    st.session_state.setup_done = True


def add_to_cart(item_row, qty):
    item_name = item_row["item"]
    category = item_row["category"]
    price = float(item_row["price"])
    eqp_type = item_row["eqp_type"]

    for cart_item in st.session_state.cart:
        if (
            cart_item["item"] == item_name
            and cart_item["category"] == category
            and cart_item["price"] == price
            and cart_item["eqp_type"] == eqp_type
        ):
            cart_item["qty"] += qty
            return

    st.session_state.cart.append({
        "item": item_name,
        "category": category,
        "price": price,
        "qty": qty,
        "eqp_type": eqp_type
    })


def get_cart_df():
    if not st.session_state.cart:
        return pd.DataFrame(columns=[
            "item", "category", "eqp_type", "price", "qty", "line_total"
        ])

    df = pd.DataFrame(st.session_state.cart)
    df["line_total"] = df["price"] * df["qty"]
    return df[["item", "category", "eqp_type", "price", "qty", "line_total"]]


def to_excel_bytes(df):
    export_df = df.rename(columns={
        "item": "Item",
        "category": "Category",
        "eqp_type": "Equipment category",
        "price": "Unit price",
        "qty": "Quantity",
        "line_total": "Line total"
    })

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Cart")
    output.seek(0)
    return output.getvalue()


def reset_setup(keep_cart=False, autoload_enabled=None):
    st.session_state.setup_done = False
    st.session_state.file_bytes = None
    st.session_state.data = None
    st.session_state.selected_sheets = []
    st.session_state.item_col = None
    st.session_state.category_col = None
    st.session_state.price_col = None

    if autoload_enabled is not None:
        st.session_state.autoload_enabled = autoload_enabled

    if not keep_cart:
        st.session_state.cart = []


def render_header():
    logo_base64 = load_logo_base64("main_logo.png")

    if logo_base64:
        st.markdown(
            f"""
            <div class="pg-header-wrap">
                <img src="data:image/png;base64,{logo_base64}" width="88">
                <div>
                    <h1 class="pg-header-title">PharmaGroup catalogue</h1>
                    <p class="pg-header-subtitle">
                        Load your default Excel from SharePoint or upload a file manually.
                    </p>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
    else:
        st.title("PharmaGroup catalogue")
        st.caption("Load your default Excel from SharePoint or upload a file manually.")


def render_available_items_desktop(filtered_df):
    st.markdown("## Available items")

    if filtered_df.empty:
        st.warning("No items match the current filters.")
        return

    h1, h2, h3, h4, h5, h6 = st.columns([4, 2, 2, 1.3, 1.2, 1.2])
    h1.markdown("**Item**")
    h2.markdown("**Equipment category**")
    h3.markdown("**Category**")
    h4.markdown("**Unit price**")
    h5.markdown("**Quantity**")
    h6.markdown("**Action**")

    st.divider()

    for idx, row in filtered_df.iterrows():
        c1, c2, c3, c4, c5, c6 = st.columns([4, 2, 2, 1.3, 1.2, 1.2])

        with c1:
            st.markdown(f"**{row['item']}**")

        with c2:
            st.write(row["eqp_type"])

        with c3:
            st.write(row["category"])

        with c4:
            st.write(f"{row['price']:.2f}")

        with c5:
            qty = st.number_input(
                "Quantity",
                min_value=1,
                max_value=1000,
                value=1,
                step=1,
                key=f"qty_desktop_{idx}",
                label_visibility="collapsed"
            )

        with c6:
            if st.button("Add", key=f"add_desktop_{idx}", use_container_width=True):
                add_to_cart(row, qty)
                st.rerun()


def render_available_items_mobile(filtered_df):
    st.markdown("## Available items")

    if filtered_df.empty:
        st.warning("No items match the current filters.")
        return

    for idx, row in filtered_df.iterrows():
        with st.container(border=True):
            st.markdown(f"<div class='pg-card-title'>{row['item']}</div>", unsafe_allow_html=True)
            st.markdown(f"<div class='pg-muted'>Equipment category: {row['eqp_type']}</div>", unsafe_allow_html=True)
            st.markdown(f"<div class='pg-muted'>Category: {row['category']}</div>", unsafe_allow_html=True)
            st.markdown(f"<div class='pg-muted'>Unit price: {row['price']:.2f}</div>", unsafe_allow_html=True)

            q_col, a_col = st.columns([1, 1])

            with q_col:
                qty = st.number_input(
                    f"Quantity {idx}",
                    min_value=1,
                    max_value=1000,
                    value=1,
                    step=1,
                    key=f"qty_mobile_{idx}"
                )

            with a_col:
                st.write("")
                if st.button("Add to cart", key=f"add_mobile_{idx}", use_container_width=True):
                    add_to_cart(row, qty)
                    st.rerun()


def render_cart_desktop(cdf):
    st.markdown("## Cart")

    if cdf.empty:
        st.info("Cart is empty.")
        return

    h1, h2, h3, h4, h5 = st.columns([3, 2, 2, 2, 1.2])
    h1.markdown("**Item**")
    h2.markdown("**Equipment category**")
    h3.markdown("**Unit price**")
    h4.markdown("**Quantity**")
    h5.markdown("**Action**")

    st.divider()

    for i, row in cdf.iterrows():
        a, b, c, d, e = st.columns([3, 1.8, 1.2, 1, 1.2])

        with a:
            st.markdown(f"**{row['item']}**")
            st.caption(row["category"])

        with b:
            st.write(row["eqp_type"])

        with c:
            st.write(f"{row['price']:.2f}")

        with d:
            st.write(f"{int(row['qty'])}")

        with e:
            if st.button("Remove", key=f"remove_desktop_{i}", use_container_width=True):
                st.session_state.cart.pop(i)
                st.rerun()


def render_cart_mobile(cdf):
    st.markdown("## Cart")

    if cdf.empty:
        st.info("Cart is empty.")
        return

    for i, row in cdf.iterrows():
        with st.container(border=True):
            st.markdown(f"<div class='pg-card-title'>{row['item']}</div>", unsafe_allow_html=True)
            st.markdown(f"<div class='pg-muted'>Category: {row['category']}</div>", unsafe_allow_html=True)
            st.markdown(f"<div class='pg-muted'>Equipment category: {row['eqp_type']}</div>", unsafe_allow_html=True)
            st.markdown(f"<div class='pg-muted'>Unit price: {row['price']:.2f}</div>", unsafe_allow_html=True)
            st.markdown(f"<div class='pg-muted'>Quantity: {int(row['qty'])}</div>", unsafe_allow_html=True)
            st.markdown(f"<div class='pg-muted'>Line total: {row['line_total']:.2f}</div>", unsafe_allow_html=True)

            if st.button("Remove", key=f"remove_mobile_{i}", use_container_width=True):
                st.session_state.cart.pop(i)
                st.rerun()


# =========================================================
# Title
# =========================================================
render_header()

# =========================================================
# Auto-load SharePoint file on first load (unless manual mode requested)
# =========================================================
if not st.session_state.setup_done and st.session_state.autoload_enabled:
    try:
        try_autoload_default_excel()
        st.rerun()
    except Exception as e:
        st.warning(f"Automatic private Excel load failed: {e}")

# =========================================================
# Setup section
# =========================================================
if not st.session_state.setup_done:
    top_a, top_b = st.columns([3, 2])

    with top_a:
        st.subheader("Setup")

    with top_b:
        if st.button("Load default SharePoint file", use_container_width=True):
            st.session_state.autoload_enabled = True
            st.session_state.setup_done = False
            st.rerun()

    uploaded_file = st.file_uploader(
        "Upload Excel file",
        type=["xlsx", "xls"],
        help="Upload an Excel workbook with one or more sheets."
    )

    if uploaded_file is None:
        st.info("Upload an Excel file to continue, or use the default SharePoint file.")
        st.stop()

    file_bytes = uploaded_file.getvalue()

    try:
        sheet_names = get_excel_sheets(file_bytes)
    except Exception as e:
        st.error(f"Could not read Excel file: {e}")
        st.stop()

    selected_sheets = st.multiselect(
        "Sheets to include",
        options=sheet_names,
        default=sheet_names
    )

    if not selected_sheets:
        st.warning("Select at least one sheet.")
        st.stop()

    try:
        raw_df = load_selected_sheets(file_bytes, selected_sheets)
    except Exception as e:
        st.error(f"Failed to load selected sheets: {e}")
        st.stop()

    if raw_df.empty:
        st.warning("No data found in selected sheets.")
        st.stop()

    available_cols = [c for c in raw_df.columns if c != "__sheet__"]

    if not available_cols:
        st.error("No usable columns found in selected sheets.")
        st.stop()

    st.markdown("### Map your columns")

    item_guess = guess_column(
        available_cols,
        [
            "polozky", "položky",
            "item", "name", "product", "part", "part name", "material", "description"
        ]
    )
    category_guess = guess_column(
        available_cols,
        [
            "kategoria", "kategória",
            "category", "group", "type", "family", "class"
        ]
    )
    price_guess = guess_column(
        available_cols,
        [
            "cena",
            "price", "cost", "unit price", "amount", "value"
        ]
    )

    c1, c2, c3 = st.columns(3)

    with c1:
        item_col = st.selectbox(
            "Item column",
            options=available_cols,
            index=available_cols.index(item_guess) if item_guess in available_cols else 0
        )

    with c2:
        category_col = st.selectbox(
            "Category column",
            options=available_cols,
            index=available_cols.index(category_guess) if category_guess in available_cols else 0
        )

    with c3:
        price_col = st.selectbox(
            "Price column",
            options=available_cols,
            index=available_cols.index(price_guess) if price_guess in available_cols else 0
        )

    with st.expander("Preview raw uploaded data"):
        st.dataframe(raw_df, use_container_width=True)

    if st.button("Confirm setup", type="primary"):
        try:
            data = build_standard_table(raw_df, item_col, category_col, price_col)
        except Exception as e:
            st.error(f"Column mapping failed: {e}")
            st.stop()

        if data.empty:
            st.warning("No valid rows after applying selected columns.")
            st.stop()

        st.session_state.file_bytes = file_bytes
        st.session_state.data = data
        st.session_state.selected_sheets = selected_sheets
        st.session_state.item_col = item_col
        st.session_state.category_col = category_col
        st.session_state.price_col = price_col
        st.session_state.setup_done = True
        st.session_state.autoload_enabled = False

        st.rerun()

    st.stop()

# =========================================================
# Compact mode (after setup)
# =========================================================
data = st.session_state.data

if data is None or data.empty:
    st.error("No data available. Please reset and load the file again.")
    if st.button("Reset", use_container_width=True):
        reset_setup(keep_cart=False, autoload_enabled=True)
        st.rerun()
    st.stop()

# Top controls
top_left, top_mid, top_right = st.columns([3, 3, 2])

with top_left:
    st.success("Excel loaded")

with top_mid:
    st.caption(f"Sheets: {', '.join(st.session_state.selected_sheets)}")

with top_right:
    if st.button("Change file / remap", use_container_width=True):
        reset_setup(keep_cart=False, autoload_enabled=False)
        st.rerun()

# Toggle
view_col1, view_col2 = st.columns([2, 2])
with view_col1:
    mobile_view = st.toggle("Mobile-friendly layout", value=True)
with view_col2:
    if st.button("Reload default SharePoint file", use_container_width=True):
        reset_setup(keep_cart=False, autoload_enabled=True)
        st.rerun()

# Filters
filter1, filter2, filter3 = st.columns([2, 2, 3])

with filter1:
    eqp_options = ["All"] + sorted(data["eqp_type"].dropna().unique().tolist())
    selected_eqp = st.selectbox("Equipment type", eqp_options)

filtered = data.copy()
if selected_eqp != "All":
    filtered = filtered[filtered["eqp_type"] == selected_eqp].copy()

with filter2:
    category_options = ["All"] + sorted(filtered["category"].dropna().unique().tolist())
    selected_category = st.selectbox("Category", category_options)

if selected_category != "All":
    filtered = filtered[filtered["category"] == selected_category].copy()

with filter3:
    search_text = st.text_input("Search item")

if search_text:
    filtered = filtered[
        filtered["item"].str.contains(search_text, case=False, na=False)
    ].copy()

filtered = filtered.sort_values(["category", "item"]).reset_index(drop=True)

cdf = get_cart_df()

# Layout
if mobile_view:
    # STACKED MOBILE LAYOUT
    render_available_items_mobile(filtered)
    st.divider()
    render_cart_mobile(cdf)
else:
    # DESKTOP LAYOUT
    left_col, right_col = st.columns([2, 1])

    with left_col:
        render_available_items_desktop(filtered)

    with right_col:
        render_cart_desktop(cdf)

# Bottom actions
st.divider()
bottom1, bottom2 = st.columns([1, 1])

with bottom1:
    if st.button("Clear cart", use_container_width=True):
        st.session_state.cart = []
        st.rerun()

with bottom2:
    excel_data = to_excel_bytes(cdf)
    st.download_button(
        label="Download Excel",
        data=excel_data,
        file_name="cart.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# Optional preview
with st.expander("Preview normalized data used by app"):
    st.dataframe(data, use_container_width=True)
import base64
from io import BytesIO
import requests
import unicodedata
from pathlib import Path