import streamlit as st
import pandas as pd
from io import BytesIO
import base64

def load_logo(path):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()


st.set_page_config(page_title="Excel Cart", layout="wide")

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

st.markdown("""
<style>
.block-container {padding-top: 1rem; padding-bottom: 1rem;}
div.stButton > button {padding-top: 0.2rem; padding-bottom: 0.2rem; min-height: 2rem;}
div[data-testid="stVerticalBlock"] > div {gap: 0.25rem;}
</style>
""", unsafe_allow_html=True)

# =========================================================
# Helpers
# =========================================================
def guess_column(columns, candidates):
    """
    Try to guess a matching column from provided candidates.
    """
    normalized = {str(c).strip().lower(): c for c in columns}
    for cand in candidates:
        if cand in normalized:
            return normalized[cand]
    return None


def get_excel_sheets(file_bytes):
    """
    Return sheet names from uploaded Excel bytes.
    """
    xl = pd.ExcelFile(BytesIO(file_bytes), engine="openpyxl")
    return xl.sheet_names


def load_selected_sheets(file_bytes, selected_sheets):
    """
    Load selected sheets and append a source sheet column.
    """
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
    Normalize user-selected columns into:
    item, category, price, sheet
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

    out = out.rename(columns={"__sheet__": "sheet"})
    out = out.reset_index(drop=True)

    return out


def add_to_cart(item_row, qty):
    """
    Add item to cart. If item already exists, increase quantity.
    """
    item_name = item_row["item"]
    category = item_row["category"]
    price = float(item_row["price"])
    sheet = item_row["sheet"]

    for cart_item in st.session_state.cart:
        if (
            cart_item["item"] == item_name and
            cart_item["category"] == category and
            cart_item["price"] == price and
            cart_item["sheet"] == sheet
        ):
            cart_item["qty"] += qty
            return

    st.session_state.cart.append({
        "item": item_name,
        "category": category,
        "price": price,
        "qty": qty,
        "sheet": sheet
    })


def get_cart_df():
    """
    Convert cart from session state to DataFrame.
    """
    if not st.session_state.cart:
        return pd.DataFrame(columns=[
            "item", "category", "sheet", "price", "qty", "line_total"
        ])

    df = pd.DataFrame(st.session_state.cart)
    df["line_total"] = df["price"] * df["qty"]

    return df[["item", "category", "sheet", "price", "qty", "line_total"]]


def reset_setup(keep_cart=False):
    """
    Reset uploaded file / mapping setup.
    """
    st.session_state.setup_done = False
    st.session_state.file_bytes = None
    st.session_state.data = None
    st.session_state.selected_sheets = []
    st.session_state.item_col = None
    st.session_state.category_col = None
    st.session_state.price_col = None

    if not keep_cart:
        st.session_state.cart = []


# =========================================================
# Title
# =========================================================

logo_base64 = load_logo("main_logo.png")

st.markdown(
    f"""
    <div style="display:flex;align-items:center;gap:15px;">
        <img src="data:image/png;base64,{logo_base64}" width="100">
        <div>
            <h1 style="margin-bottom:0;">PharmaGroup catalogue</h1>
            <p style="margin-top:0;color:gray;">
                Upload an Excel file, map columns once, then use compact filters and a cart.
            </p>
        </div>
    </div>
    """,
    unsafe_allow_html=True
)



# =========================================================
# Setup section (visible only before confirm)
# =========================================================
if not st.session_state.setup_done:
    st.subheader("Setup")

    uploaded_file = st.file_uploader(
        "Upload Excel file",
        type=["xlsx", "xls"],
        help="Upload an Excel workbook with one or more sheets."
    )

    if uploaded_file is None:
        st.info("Upload an Excel file to start.")
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
        ["item", "name", "product", "part", "part name", "material", "description"]
    )
    category_guess = guess_column(
        available_cols,
        ["category", "group", "type", "family", "class"]
    )
    price_guess = guess_column(
        available_cols,
        ["price", "cost", "unit price", "amount", "value"]
    )

    c1, c2, c3 = st.columns(3)

    with c1:
        item_col = st.selectbox(
            "Item column",
            options=available_cols,
            index=available_cols.index(item_guess) if item_guess in available_cols else 0,
            help="Column containing item / product / part name."
        )

    with c2:
        category_col = st.selectbox(
            "Category column",
            options=available_cols,
            index=available_cols.index(category_guess) if category_guess in available_cols else 0,
            help="Column containing category / group / family."
        )

    with c3:
        price_col = st.selectbox(
            "Price column",
            options=available_cols,
            index=available_cols.index(price_guess) if price_guess in available_cols else 0,
            help="Column containing numeric item price."
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

        st.rerun()

    st.stop()

# =========================================================
# Compact mode (after setup)
# =========================================================
data = st.session_state.data

if data is None or data.empty:
    st.error("No data available. Please reset and load the file again.")
    if st.button("Reset"):
        reset_setup(keep_cart=False)
        st.rerun()
    st.stop()

# Top compact controls
top_left, top_mid, top_right = st.columns([3, 2, 1])

with top_left:
    st.success("Excel loaded")

with top_mid:
    file_info = f"Sheets: {', '.join(st.session_state.selected_sheets)}"
    st.caption(file_info)

with top_right:
    if st.button("Change file / remap"):
        reset_setup(keep_cart=False)
        st.rerun()

# =========================================================
# Compact dropdown filters
# =========================================================
filter1, filter2, filter3 = st.columns([2, 2, 3])

with filter1:
    sheet_options = ["All"] + sorted(data["sheet"].dropna().unique().tolist())
    selected_sheet = st.selectbox("Sheet", sheet_options)

filtered = data.copy()
if selected_sheet != "All":
    filtered = filtered[filtered["sheet"] == selected_sheet].copy()

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

# =========================================================
# Main layout
# =========================================================
left_col, right_col = st.columns([2, 1])

# ---------------------------------------------------------
# Available items
# ---------------------------------------------------------
with left_col:
    st.markdown("## Available items")

    if filtered.empty:
        st.warning("No items match the current filters.")
    else:
        # Header row
        h1, h2, h3, h4, h5 = st.columns([3.5, 2, 1.3, 1.2, 1.2])
        with h1:
            st.markdown("**Item**")
        with h2:
            st.markdown("**Category**")
        with h3:
            st.markdown("**Unit price (€)**")
        with h4:
            st.markdown("**Quantity**")
        with h5:
            st.markdown("**Action**")

        st.divider()

        # Data rows
        for idx, row in filtered.iterrows():
            c1, c2, c3, c4, c5 = st.columns([3.5, 2, 1.3, 1.2, 1.2])

            with c1:
                st.markdown(
                    f"""
                    <div style="line-height:1.1;">
                        <div style="font-weight:600;">{row['item']}</div>
                        <div style="font-size:12px;color:gray;">Sheet: {row['sheet']}</div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

            with c2:
                st.write(row["category"])

            with c3:
                st.write(f"{row['price']:.2f}")

            with c4:
                qty = st.number_input(
                    "Quantity",
                    min_value=1,
                    max_value=1000,
                    value=1,
                    step=1,
                    key=f"qty_{idx}",
                    label_visibility="collapsed"
                )

            with c5:
                if st.button("Add", key=f"add_{idx}", use_container_width=True):
                    add_to_cart(row, qty)
                    st.rerun()
# ---------------------------------------------------------
# Cart
# ---------------------------------------------------------
with right_col:
    st.markdown("## Cart")

    cdf = get_cart_df()

    if cdf.empty:
        st.info("Cart is empty.")
    else:
        for i, row in cdf.iterrows():
            a, b, c = st.columns([3, 1.2, 1])

            with a:
                st.write(f"**{row['item']}**")
                st.caption(f"{row['category']} | {row['sheet']}")

            with b:
                st.write(f"{int(row['qty'])} × {row['price']:.2f}")

            with c:
                if st.button("Remove", key=f"remove_{i}"):
                    st.session_state.cart.pop(i)
                    st.rerun()

        st.divider()
        total = cdf["line_total"].sum()
        st.markdown(f"## Total: {total:.2f}")

        btn1, btn2 = st.columns(2)

        with btn1:
            if st.button("Clear cart"):
                st.session_state.cart = []
                st.rerun()

        with btn2:
            st.download_button(
                label="Download CSV",
                data=cdf.to_csv(index=False).encode("utf-8"),
                file_name="cart.csv",
                mime="text/csv"
            )

# =========================================================
# Optional preview
# =========================================================
with st.expander("Preview normalized data used by app"):
    st.dataframe(data, use_container_width=True)
