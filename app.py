import io
import re

import pandas as pd
import streamlit as st

st.set_page_config(page_title="ระบบรวมยอดขาย Multi-Platform", layout="wide", page_icon="📦")

st.title("📦 ระบบรวมยอดขาย Shopee / Lazada / TikTok")
st.subheader("แปลงและจัดกลุ่มประเภทสินค้า (Neo / DN) ให้อยู่ในไฟล์ Excel เดียวกัน")

TARGET_COLUMNS = [
    "วันที่ทำการสั่งซื้อ", "order number", "ชื่อผู้รับ", "ที่อยู่ในการจัดส่ง",
    "ประเภท(Neo,DN)", "ชื่อสินค้า", "SKU", "ตัวเลือก", "จำนวน",
    "ราคาขายสุทธิ", "ค่าส่ง", "ช่องทางการชำระเงิน", "ตัวเลือกการจัดส่ง",
    "หมายเลขติดตามพัสดุ",
]

# ---------------------------------------------------------
# Mapping คอลัมน์ของแต่ละแพลตฟอร์ม (เรียงตามลำดับความสำคัญ)
# ---------------------------------------------------------
MAPPING = {
    "Shopee": {
        "date": ["เวลาการชำระเงิน", "เวลาที่สั่งซื้อ", "Order Creation Date"],
        "order_id": ["หมายเลขคำสั่งซื้อ", "Order ID"],
        "recipient": ["ชื่อผู้รับ", "Receiver Name"],
        "address": ["ที่อยู่ในการจัดส่ง", "Shipping Address"],
        "product_name": ["ชื่อสินค้า", "Product Name"],
        "sku": ["เลขอ้างอิง SKU (SKU Reference No.)", "เลขชี้วัด SKU", "SKU Reference No.", "SKU"],
        "variation": ["ชื่อตัวเลือก", "Variation Name"],
        "qty": ["จำนวน", "Quantity"],
        "price": ["ราคาขายสุทธิ", "ราคาขายต่อชิ้น", "Deal Price", "ราคาตั้งต้น"],
        "shipping_fee": ["ค่าจัดส่งที่ชำระโดยผู้ซื้อ", "Estimated Shipping Fee"],
        "payment": ["ช่องทางการชำระเงิน", "Payment Method"],
        "courier": ["ตัวเลือกการจัดส่ง", "Shipping Option"],
        "tracking": ["*หมายเลขติดตามพัสดุ", "หมายเลขติดตามพัสดุ", "Tracking Number"],
    },
    "Lazada": {
        "date": ["createTime", "Created Time", "Order Date"],
        "order_id": ["orderNumber", "Order ID"],
        "recipient": ["customerName", "Recipient Name"],
        "address": ["shippingAddress", "Address"],
        "product_name": ["itemName", "Item Name"],
        "sku": ["sellerSku", "Seller SKU", "SKU"],
        "variation": ["variation", "Variation"],
        "qty": ["quantity", "Quantity"],
        "price": ["paidPrice", "unitPrice", "Item Price"],
        "shipping_fee": ["shippingFee", "Shipping Fee"],
        "payment": ["payMethod", "Payment Method"],
        "courier": ["shippingProvider", "Shipment Provider"],
        "tracking": ["trackingCode", "Tracking Code"],
    },
    "TikTok": {
        "date": ["Created Time", "Order Time"],
        "order_id": ["Order ID"],
        "recipient": ["Recipient", "Customer Name"],
        "address": ["Full Address", "Detail Address"],
        "product_name": ["Product Name"],
        "sku": ["Seller SKU", "SKU ID"],
        "variation": ["Variation"],
        "qty": ["Quantity"],
        "price": ["SKU Subtotal After Discount", "SKU Unit Price"],
        "shipping_fee": ["Shipping Fee", "Customer Paid Shipping Fee"],
        "payment": ["Payment Method"],
        "courier": ["Shipping Provider", "Delivery Option"],
        "tracking": ["Tracking ID", "Tracking Number"],
    },
}


# ---------------------------------------------------------
# Helper
# ---------------------------------------------------------
def norm_sku(value):
    """ทำ SKU ให้อยู่ในรูปมาตรฐาน เพื่อเทียบกับ Master ได้แม่นยำ"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    s = str(value).strip()
    if s.lower() == "nan":
        return ""
    s = re.sub(r"\.0$", "", s)  # 123.0 -> 123
    return s.upper()


def find_col(df, candidates):
    lookup = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        key = cand.strip().lower()
        if key in lookup:
            return lookup[key]
    return None


def get_series(df, candidates):
    col = find_col(df, candidates)
    if col is None:
        return pd.Series("", index=df.index, dtype=object), None
    return df[col], str(col).strip().lower()


def to_num(series):
    cleaned = series.astype(str).str.replace(",", "", regex=False).str.strip()
    return pd.to_numeric(cleaned, errors="coerce")


def detect_platform(df):
    cols = {str(c).strip().lower() for c in df.columns}
    if "order id" in cols and ("sku subtotal after discount" in cols or "seller sku" in cols):
        return "TikTok"
    if cols & {"ordernumber", "sellersku", "createtime"}:
        return "Lazada"
    if cols & {"หมายเลขคำสั่งซื้อ", "เวลาการชำระเงิน", "เวลาที่สั่งซื้อ"}:
        return "Shopee"
    return "Unknown"


def read_any(file):
    name = file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(file, dtype=str, encoding="utf-8-sig")
    return pd.read_excel(file, dtype=str)


# ---------------------------------------------------------
# 1. โหลด SKU จาก Master Data
# ---------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_master_skus(master_bytes):
    neo_skus, dn_skus = set(), set()
    xls = pd.ExcelFile(io.BytesIO(master_bytes))

    for sheet in xls.sheet_names:
        df_sheet = pd.read_excel(xls, sheet_name=sheet, dtype=str)
        tokens = set(re.split(r"[^a-z0-9]+", str(sheet).strip().lower()))
        sheet_is_neo = "neo" in tokens
        sheet_is_dn = "dn" in tokens

        for col in df_sheet.columns:
            col_clean = str(col).strip().lower()
            col_tokens = set(re.split(r"[^a-z0-9]+", col_clean))
            has_sku = "sku" in col_clean

            target = None
            # ชื่อคอลัมน์ระบุชัดเจนก่อน เช่น "Neo SKU", "DN SKU", "Neo", "DN"
            if "neo" in col_tokens:
                target = neo_skus
            elif "dn" in col_tokens:
                target = dn_skus
            # ถ้าไม่ระบุ ให้ใช้ชื่อ Sheet (เฉพาะคอลัมน์ที่เป็น SKU)
            elif has_sku and sheet_is_neo and not sheet_is_dn:
                target = neo_skus
            elif has_sku and sheet_is_dn and not sheet_is_neo:
                target = dn_skus

            if target is not None:
                values = df_sheet[col].dropna().map(norm_sku)
                target.update(v for v in values if v)

    return neo_skus, dn_skus


# ---------------------------------------------------------
# 2. ประมวลผลไฟล์ยอดขาย
# ---------------------------------------------------------
def process_sales_file(file, neo_skus, dn_skus, deduplicate_shipping=True):
    df = read_any(file)
    df.columns = [str(c).strip() for c in df.columns]

    platform = detect_platform(df)
    if platform == "Unknown":
        return None, platform

    m = MAPPING[platform]
    out = pd.DataFrame(index=df.index)

    out["วันที่ทำการสั่งซื้อ"] = get_series(df, m["date"])[0]
    out["order number"] = get_series(df, m["order_id"])[0].fillna("").astype(str).str.strip()
    out["ชื่อผู้รับ"] = get_series(df, m["recipient"])[0]
    out["ที่อยู่ในการจัดส่ง"] = get_series(df, m["address"])[0]
    out["ชื่อสินค้า"] = get_series(df, m["product_name"])[0]
    out["SKU"] = get_series(df, m["sku"])[0].map(norm_sku)
    out["ตัวเลือก"] = get_series(df, m["variation"])[0]

    qty = to_num(get_series(df, m["qty"])[0]).fillna(1)
    out["จำนวน"] = qty

    price_raw, price_col = get_series(df, m["price"])
    price = to_num(price_raw).fillna(0)
    # TikTok: SKU Subtotal After Discount เป็นยอดรวมของแถว ต้องหารจำนวนเพื่อให้เป็นราคาต่อชิ้น
    if platform == "TikTok" and price_col == "sku subtotal after discount":
        price = price / qty.replace(0, 1)
    out["ราคาขายสุทธิ"] = price

    out["ค่าส่ง"] = to_num(get_series(df, m["shipping_fee"])[0]).fillna(0)
    out["ช่องทางการชำระเงิน"] = get_series(df, m["payment"])[0]
    out["ตัวเลือกการจัดส่ง"] = get_series(df, m["courier"])[0]
    out["หมายเลขติดตามพัสดุ"] = get_series(df, m["tracking"])[0]

    # ตัดแถวที่ไม่ใช่ออเดอร์จริง (เช่น แถวคำอธิบายใต้หัวตารางของ TikTok / แถวว่าง)
    valid = out["order number"].ne("") & out["order number"].str.lower().ne("nan")
    if platform == "TikTok":
        valid &= out["order number"].str.fullmatch(r"\d+")
    out = out[valid].copy()

    # ป้องกันค่าส่งซ้ำในออเดอร์เดียวกัน
    if deduplicate_shipping:
        out.loc[out.duplicated("order number", keep="first"), "ค่าส่ง"] = 0

    def categorize(sku):
        if not sku:
            return ""
        if sku in neo_skus:
            return "Neo"
        if sku in dn_skus:
            return "DN"
        return ""

    out["ประเภท(Neo,DN)"] = out["SKU"].map(categorize)

    out = out[TARGET_COLUMNS].fillna("")
    out["แพลตฟอร์ม"] = platform
    out["ยอดรวมสินค้า"] = out["ราคาขายสุทธิ"] * out["จำนวน"]
    return out.reset_index(drop=True), platform


# ---------------------------------------------------------
# 3. UI
# ---------------------------------------------------------
st.sidebar.header("⚙️ การตั้งค่า & ไฟล์อ้างอิง")

master_file = st.sidebar.file_uploader(
    "1. อัปโหลดไฟล์ Master Data (Excel)",
    type=["xlsx", "xls"],
    help="ไฟล์ Excel ที่มีแท็บหรือคอลัมน์ 'Neo SKU' และ 'DN SKU'",
)

dedup_shipping = st.sidebar.checkbox(
    "แสดงค่าส่งเฉพาะแถวแรกของออเดอร์ (ป้องกันค่าส่งซ้ำ)", value=True
)

st.divider()

sales_files = st.file_uploader(
    "2. ลากไฟล์ยอดขายจาก Shopee, Lazada, TikTok มาวางที่นี่ (รองรับหลายไฟล์)",
    type=["csv", "xlsx", "xls"],
    accept_multiple_files=True,
)

if sales_files:
    neo_skus, dn_skus = set(), set()
    if master_file:
        try:
            neo_skus, dn_skus = load_master_skus(master_file.getvalue())
            st.success(
                f"โหลด Master Data เรียบร้อย! (Neo SKU: {len(neo_skus):,} รายการ / DN SKU: {len(dn_skus):,} รายการ)"
            )
            both = neo_skus & dn_skus
            if both:
                st.warning(f"พบ SKU ที่อยู่ทั้ง Neo และ DN จำนวน {len(both)} รายการ (ระบบจะจัดเป็น Neo)")
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการอ่านไฟล์ Master Data: {e}")
    else:
        st.warning("⚠️ ยังไม่ได้อัปโหลด Master Data (ช่อง 'ประเภท(Neo,DN)' จะว่างไว้)")

    all_processed = []
    for f in sales_files:
        try:
            processed_df, plat = process_sales_file(f, neo_skus, dn_skus, dedup_shipping)
        except Exception as e:
            st.error(f"❌ อ่านไฟล์ {f.name} ไม่ได้: {e}")
            continue
        if processed_df is None:
            st.error(f"❌ ไฟล์ {f.name}: ไม่สามารถระบุแพลตฟอร์มได้ (ตรวจสอบหัวคอลัมน์)")
            continue
        all_processed.append(processed_df)
        st.info(f"📄 **{f.name}** → แพลตฟอร์ม: **{plat}** ({len(processed_df):,} แถว)")

    if all_processed:
        final_df = pd.concat(all_processed, ignore_index=True)

        st.divider()
        st.subheader("📊 สรุปผลลัพธ์ข้อมูลยอดขาย")

        is_neo = final_df["ประเภท(Neo,DN)"] == "Neo"
        is_dn = final_df["ประเภท(Neo,DN)"] == "DN"
        unmatched = final_df["ประเภท(Neo,DN)"] == ""

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("รายการทั้งหมด", f"{len(final_df):,} แถว")
        c2.metric("Neo", f"{is_neo.sum():,} รายการ", f"฿{final_df.loc[is_neo, 'ยอดรวมสินค้า'].sum():,.2f}", delta_color="off")
        c3.metric("DN", f"{is_dn.sum():,} รายการ", f"฿{final_df.loc[is_dn, 'ยอดรวมสินค้า'].sum():,.2f}", delta_color="off")
        c4.metric("ยังจับคู่ไม่ได้", f"{unmatched.sum():,} รายการ")

        # ตารางสรุปตามแพลตฟอร์ม x ประเภท
        summary = (
            final_df.assign(**{"ประเภท": final_df["ประเภท(Neo,DN)"].replace("", "ไม่ระบุ")})
            .groupby(["แพลตฟอร์ม", "ประเภท"], as_index=False)
            .agg(จำนวนแถว=("order number", "size"), จำนวนชิ้น=("จำนวน", "sum"), ยอดรวมสินค้า=("ยอดรวมสินค้า", "sum"))
        )
        st.markdown("**สรุปตามแพลตฟอร์มและประเภท**")
        st.dataframe(summary, use_container_width=True, hide_index=True)

        # SKU ที่ยังจับคู่ไม่ได้
        unmatched_skus = (
            final_df.loc[unmatched & (final_df["SKU"] != ""), ["แพลตฟอร์ม", "SKU", "ชื่อสินค้า"]]
            .drop_duplicates("SKU")
            .reset_index(drop=True)
        )
        if not unmatched_skus.empty:
            with st.expander(f"🔍 SKU ที่ไม่พบใน Master Data ({len(unmatched_skus)} รายการ)"):
                st.dataframe(unmatched_skus, use_container_width=True, hide_index=True)

        # ตัวกรอง
        f1, f2 = st.columns(2)
        plat_filter = f1.multiselect("กรองแพลตฟอร์ม", sorted(final_df["แพลตฟอร์ม"].unique()))
        type_filter = f2.multiselect("กรองประเภท", ["Neo", "DN", ""], format_func=lambda x: x or "ยังไม่ระบุ")

        view_df = final_df
        if plat_filter:
            view_df = view_df[view_df["แพลตฟอร์ม"].isin(plat_filter)]
        if type_filter:
            view_df = view_df[view_df["ประเภท(Neo,DN)"].isin(type_filter)]

        st.dataframe(view_df, use_container_width=True)

        # ไฟล์ Excel สำหรับดาวน์โหลด
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            final_df.to_excel(writer, index=False, sheet_name="Sales_Report")
            summary.to_excel(writer, index=False, sheet_name="Summary")
            if not unmatched_skus.empty:
                unmatched_skus.to_excel(writer, index=False, sheet_name="Unmatched_SKU")
            for ws in writer.book.worksheets:
                for col_cells in ws.columns:
                    width = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells[:200])
                    ws.column_dimensions[col_cells[0].column_letter].width = min(max(width + 2, 10), 50)
        output.seek(0)

        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel รวมยอดขาย",
            data=output,
            file_name="รวมยอดขาย_Shopee_Lazada_TikTok.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("👈 อัปโหลด Master Data ที่แถบด้านซ้าย แล้วลากไฟล์ยอดขายมาวางด้านบนได้เลย")
