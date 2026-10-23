import os
import re
import io
import json
import sqlite3
import zipfile
from datetime import datetime, timedelta
import pandas as pd
from PIL import Image
import streamlit as st

# ==================== CẤU HÌNH TRANG WEB STREAMLIT ====================
ICON_PATH = "cloudbi-icon-7156.png"
page_icon = ICON_PATH if os.path.exists(ICON_PATH) else "📊"

st.set_page_config(
    page_title="QUẢN LÝ BÁO CÁO KIỂM VẢI · 验布报告管理系统",
    page_icon=page_icon,
    layout="wide",
    initial_sidebar_state="expanded"
)

DB_PATH = "BaoCaoKiemVai.db"

# ==================== KHỞI TẠO CƠ SỞ DỮ LIỆU SQLITE ====================
def get_db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    return conn

def init_db():
    conn = get_db()
    cur = conn.cursor()
    cur.execute('''
        CREATE TABLE IF NOT EXISTS reports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            lot TEXT,
            style TEXT,
            color TEXT,
            customer TEXT,
            grade TEXT,
            notes TEXT,
            main_image TEXT,
            defect_images TEXT,
            defect_count INTEGER DEFAULT 0,
            upload_date TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    cur.execute('''
        CREATE TABLE IF NOT EXISTS customers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE
        )
    ''')
    cur.execute('''
        CREATE TABLE IF NOT EXISTS inventory_stock (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            customer TEXT,
            in_date TEXT,
            style TEXT,
            lot TEXT,
            color TEXT,
            customs_no TEXT,
            fabric_code TEXT,
            fabric_type TEXT,
            composition TEXT,
            weight REAL,
            rolls REAL,
            yards REAL,
            invoice TEXT,
            notes TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    cur.execute('''
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )
    ''')
    default_custs = ["BD", "CABI", "CTP", "FG", "GA", "GF", "LH", "MK MEN", "MK WOMEN"]
    for c in default_custs:
        cur.execute("INSERT OR IGNORE INTO customers (name) VALUES (?)", (c,))
    cur.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('storage_dir', ?)", ("D:/BaoCaoKiemVai/Original",))
    conn.commit()
    conn.close()

init_db()

# ==================== TIỆN ÍCH CHUẨN HÓA DỮ LIỆU ====================
def remove_vietnamese_tones(text):
    if not text:
        return ""
    text = re.sub(r'[àáạảãâầấậẩẫăằắặẳẵ]', 'a', text)
    text = re.sub(r'[ÀÁẠẢÃÂẦẤẬẨẪĂẰẮẶẲẴ]', 'A', text)
    text = re.sub(r'[èéẹẻẽêềếệểễ]', 'e', text)
    text = re.sub(r'[ÈÉẸẺẼÊỀẾỆỂỄ]', 'E', text)
    text = re.sub(r'[ìíịỉĩ]', 'i', text)
    text = re.sub(r'[ÌÍỊỈĨ]', 'I', text)
    text = re.sub(r'[òóọỏõôồốộổỗơờớợởỡ]', 'o', text)
    text = re.sub(r'[ÒÓỌỎÕÔỒỐỘỔỖƠỜỚỢỞỠ]', 'O', text)
    text = re.sub(r'[ùúụủũưừứựửữ]', 'u', text)
    text = re.sub(r'[ÙÚỤỦŨƯỪỨỰỬỮ]', 'U', text)
    text = re.sub(r'[ỳýỵỷỹ]', 'y', text)
    text = re.sub(r'[ỲÝỴỶỸ]', 'Y', text)
    text = re.sub(r'[đ]', 'd', text)
    text = re.sub(r'[Đ]', 'D', text)
    return text

def sanitize_for_filename(text):
    clean = remove_vietnamese_tones(text)
    clean = re.sub(r'[\\/*?:"<>|]', '', clean)
    clean = re.sub(r'\s+', ' ', clean)
    return clean.strip()

def format_excel_date(val):
    if val is None or pd.isna(val):
        return ""
    if isinstance(val, (datetime, pd.Timestamp)):
        return val.strftime("%Y/%m/%d")
    val_str = str(val).strip()
    if not val_str:
        return ""
    try:
        val_float = float(val_str)
        if 20000 < val_float < 80000:
            converted = datetime(1899, 12, 30) + timedelta(days=val_float)
            return converted.strftime("%Y/%m/%d")
    except Exception:
        pass
    if len(val_str) == 8 and val_str.isdigit():
        return f"{val_str[:4]}/{val_str[4:6]}/{val_str[6:]}"
    clean_date = val_str.split()[0].replace("-", "/").replace(".", "/")
    parts = clean_date.split("/")
    if len(parts) == 3:
        if len(parts[0]) == 4:
            return f"{int(parts[0]):04d}/{int(parts[1]):02d}/{int(parts[2]):02d}"
        elif len(parts[2]) == 4:
            return f"{int(parts[2]):04d}/{int(parts[1]):02d}/{int(parts[0]):02d}"
    return val_str

def get_customers_list():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT DISTINCT name FROM customers WHERE TRIM(name) != '' ORDER BY name ASC")
    custs = [r[0] for r in cur.fetchall()]
    conn.close()
    return custs

def get_storage_dir_setting():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT value FROM settings WHERE key='storage_dir'")
    row = cur.fetchone()
    conn.close()
    val = row[0] if row else "D:/BaoCaoKiemVai/Original"
    val = val.replace("\\", "/").rstrip("/")
    return val

def set_storage_dir_setting(val):
    clean_val = val.strip().replace("\\", "/").rstrip("/")
    if not clean_val:
        clean_val = "D:/BaoCaoKiemVai/Original"
    try:
        os.makedirs(clean_val, exist_ok=True)
    except Exception as e:
        pass
    conn = get_db()
    cur = conn.cursor()
    cur.execute("REPLACE INTO settings (key, value) VALUES ('storage_dir', ?)", (clean_val,))
    conn.commit()
    conn.close()
    return clean_val

# Khởi tạo thư mục lưu trữ thực tế
CURRENT_STORAGE_DIR = get_storage_dir_setting()
try:
    os.makedirs(CURRENT_STORAGE_DIR, exist_ok=True)
except Exception:
    CURRENT_STORAGE_DIR = "storage"
    os.makedirs(CURRENT_STORAGE_DIR, exist_ok=True)

# Hàm gọi hộp thoại chọn thư mục Windows trực tiếp (Local Machine)
def open_native_folder_dialog(default_dir):
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        selected_dir = filedialog.askdirectory(initialdir=default_dir, title="Chọn thư mục lưu trữ báo cáo kiểm vải")
        root.destroy()
        if selected_dir:
            return selected_dir.replace("\\", "/")
    except Exception:
        pass
    return None

# ==================== TÙY CHỈNH CSS GIAO DIỆN ====================
st.markdown("""
    <style>
    .app-header {
        background-color: #0f172a;
        color: white;
        padding: 10px 18px;
        border-radius: 6px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 15px;
    }
    .app-header-title {
        font-size: 16px;
        font-weight: 700;
        letter-spacing: 0.5px;
    }
    .app-header-status {
        color: #38bdf8;
        font-size: 13px;
        font-weight: 600;
    }
    .image-preview-frame {
        border: 2px dashed #0f172a;
        background-color: #0f172a;
        border-radius: 4px;
        padding: 6px;
        text-align: center;
        min-height: 380px;
        display: flex;
        flex-direction: column;
        justify-content: flex-start;
        align-items: center;
        overflow: auto;
    }
    .dim-badge {
        background-color: #f1f5f9;
        color: #64748b;
        padding: 3px 10px;
        border-radius: 4px;
        font-size: 12px;
        font-weight: 700;
        display: inline-block;
    }
    .filename-box {
        background-color: #eff6ff;
        color: #1d4ed8;
        border: 1px solid #bfdbfe;
        border-radius: 4px;
        padding: 8px 12px;
        font-weight: 700;
        font-size: 13px;
        margin-top: 4px;
        margin-bottom: 10px;
        word-break: break-all;
    }
    .stButton button {
        border-radius: 4px;
        font-size: 13px;
        font-weight: 600;
    }
    </style>
""", unsafe_allow_html=True)

# ==================== THANH ĐIỀU HƯỚNG BÊN TRÁI ====================
with st.sidebar:
    if os.path.exists(ICON_PATH):
        st.image(ICON_PATH, width=54)
    st.markdown("### 📋 MENU QUẢN LÝ")
    st.caption("验布报告在线管理系统")
    st.divider()

    nav_option = st.radio(
        "Chọn chức năng / 功能选择:",
        [
            "➕ Thêm báo cáo / 添加报告",
            "📊 Tổng quan / 概览",
            "🔍 Tra cứu / 查询",
            "📋 Danh Sách Nhập Kho / 入库清单",
            "⚙ Cài đặt _Thư mục / 设置"
        ],
        index=0
    )
    st.divider()
    active_dir_display = get_storage_dir_setting()
    st.caption(f"📁 Thư mục lưu hiện tại:\n`{active_dir_display}`")

# Header Bar
st.markdown("""
    <div class="app-header">
        <div class="app-header-title">QUẢN LÝ BÁO CÁO KIỂM VẢI / 验布报告管理系统</div>
        <div class="app-header-status">⚡ Hệ thống sẵn sàng</div>
    </div>
""", unsafe_allow_html=True)

# Session states
if "selected_customer" not in st.session_state:
    st.session_state.selected_customer = "MK MEN"
if "selected_grade" not in st.session_state:
    st.session_state.selected_grade = "Loại A (A 级)"
if "zoom_level" not in st.session_state:
    st.session_state.zoom_level = 49

# ==============================================================================
# TAB 1: THÊM BÁO CÁO
# ==============================================================================
if nav_option == "➕ Thêm báo cáo / 添加报告":

    conn = get_db()
    df_inventory = pd.read_sql_query("SELECT lot, style, color, customer FROM inventory_stock ORDER BY id DESC", conn)
    conn.close()

    suggest_map = {}
    suggest_labels = ["-- Gõ để tìm hoặc chọn LOT từ dữ liệu nhập kho --"]
    if not df_inventory.empty:
        for _, row in df_inventory.iterrows():
            l_val = str(row["lot"] or "").strip()
            s_val = str(row["style"] or "").strip()
            c_val = str(row["color"] or "").strip()
            cust_val = str(row["customer"] or "").strip()
            if l_val:
                label = f"{l_val}  |  {s_val}  |  {c_val}  |  {cust_val}"
                suggest_labels.append(label)
                suggest_map[label] = {
                    "lot": l_val,
                    "style": s_val,
                    "color": c_val,
                    "customer": cust_val
                }

    col_left, col_right = st.columns([5, 5], gap="large")

    # ------------------ CỘT TRÁI: HÌNH ẢNH ------------------
    with col_left:
        c_title, c_dim, c_btn = st.columns([5, 2, 5])
        with c_title:
            st.markdown("<b style='font-size:14px; color:#0f172a;'>ẢNH BÁO CÁO GỐC / 原始报告图片</b>", unsafe_allow_html=True)
        with c_dim:
            dim_text = st.session_state.get("current_image_dims", "1240 × 1753px")
            st.markdown(f"<div class='dim-badge'>{dim_text}</div>", unsafe_allow_html=True)
        with c_btn:
            uploaded_main_file = st.file_uploader(
                "Chọn ảnh báo cáo",
                type=["jpg", "jpeg", "png", "bmp"],
                key="main_uploader",
                label_visibility="collapsed"
            )

        tb1, tb2, tb3, tb4, tb5, tb6 = st.columns([2.2, 2.2, 2, 2, 3.2, 2])
        with tb1:
            if st.button("📐 Nắn góc", use_container_width=True, type="secondary"):
                st.session_state.crop_warp_notice = True
        with tb2:
            if st.button("🔎 Vùng đầu", use_container_width=True, type="primary"):
                st.session_state.zoom_level = 95
        with tb3:
            if st.button("⛶ Trọn trang", use_container_width=True):
                st.session_state.zoom_level = 40
        with tb4:
            if st.button("↔ Vừa rộng", use_container_width=True):
                st.session_state.zoom_level = 65
        with tb5:
            new_zoom = st.slider(
                "Thu phóng",
                min_value=20,
                max_value=200,
                value=st.session_state.zoom_level,
                step=1,
                format="+ %d%%",
                label_visibility="collapsed"
            )
            st.session_state.zoom_level = new_zoom
        with tb6:
            if st.button("🔄 Đổi ảnh", use_container_width=True):
                st.session_state.main_image_cache = None
                st.rerun()

        if st.session_state.get("crop_warp_notice", False):
            st.info("💡 Tính năng Nắn góc: Khi tải ảnh chụp từ tài liệu, hệ thống tự động khử cong và căn chỉnh thẳng theo khổ đứng.")
            st.session_state.crop_warp_notice = False

        if uploaded_main_file is not None:
            pil_main = Image.open(uploaded_main_file)
            st.session_state["current_image_dims"] = f"{pil_main.width} × {pil_main.height}px"
            st.session_state.main_image_cache = pil_main
        elif "main_image_cache" not in st.session_state or st.session_state.main_image_cache is None:
            st.session_state["current_image_dims"] = "0 × 0px"
            st.session_state.main_image_cache = None

        if st.session_state.main_image_cache is not None:
            img = st.session_state.main_image_cache
            zoom_factor = st.session_state.zoom_level / 100.0
            display_width = int(img.width * zoom_factor * 0.45)
            display_width = max(300, min(display_width, 850))

            st.markdown("<div class='image-preview-frame'>", unsafe_allow_html=True)
            st.image(img, width=display_width)
            st.markdown("</div>", unsafe_allow_html=True)
        else:
            st.markdown("""
                <div class="image-preview-frame" style="justify-content: center;">
                    <div style="color: #94a3b8; font-weight: bold; font-size: 13px;">
                        Khu vực hiển thị ảnh báo cáo chính / 报告主图预览<br>
                        (Bấm 'Chọn ảnh báo cáo' ở trên)
                    </div>
                </div>
            """, unsafe_allow_html=True)

        st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)
        d_title, d_btn = st.columns([6, 6])
        with d_title:
            st.markdown("<b style='font-size: 13px; color: #4338ca;'>ALBUM ẢNH LỖI CỦA LOT / 疵点照片:</b>", unsafe_allow_html=True)
        with d_btn:
            uploaded_defects = st.file_uploader(
                "+ Thêm ảnh lỗi / 添加疵点图",
                type=["jpg", "jpeg", "png"],
                accept_multiple_files=True,
                key="defect_uploader",
                label_visibility="collapsed"
            )

        if uploaded_defects:
            st.markdown(f"<span style='font-size:12px; color:#64748b;'>Đã đính kèm <b>{len(uploaded_defects)}</b> ảnh lỗi</span>", unsafe_allow_html=True)
            d_cols = st.columns(min(6, len(uploaded_defects)))
            for idx, df_item in enumerate(uploaded_defects):
                with d_cols[idx % 6]:
                    st.image(Image.open(df_item), caption=f"Lỗi {idx+1}", use_column_width=True)
        else:
            st.markdown("""
                <div style="background:white; border:1px solid #cbd5e1; border-radius:4px; height:70px; display:flex; align-items:center; justify-content:center; color:#94a3b8; font-size:12px;">
                    Chưa có ảnh lỗi đính kèm
                </div>
            """, unsafe_allow_html=True)

    # ------------------ CỘT PHẢI: FORM THÔNG TIN ------------------
    with col_right:
        f_title, f_btn1, f_btn2 = st.columns([4.5, 3.8, 3.7])
        with f_title:
            st.markdown("<b style='font-size: 14px; color: #0f172a;'>THÔNG TIN BÁO CÁO / 报告信息</b>", unsafe_allow_html=True)
        with f_btn1:
            excel_file_upload = st.file_uploader(
                "📥 Nhập Excel",
                type=["xlsx", "xls"],
                key="excel_inventory_upload",
                label_visibility="collapsed"
            )
        with f_btn2:
            if st.button("📋 DS Nhập Kho", help="Mở danh sách chi tiết các dòng nhập kho"):
                st.session_state.open_inventory_modal = True

        if excel_file_upload is not None:
            try:
                df_read = pd.read_excel(excel_file_upload)
                def find_c(kw_list):
                    for col_name in df_read.columns:
                        c_str = str(col_name).lower()
                        if any(k in c_str for k in kw_list):
                            return col_name
                    return None

                c_cust = find_c(['客户', 'khách hàng', 'khach hang'])
                c_date = find_c(['入库日期', 'ngày nhập kho', 'ngay nhap'])
                c_style = find_c(['款号', 'mã hàng', 'ma hang'])
                c_lot = find_c(['缸号', 'lot'])
                c_color = find_c(['颜色', 'màu sắc', 'mau sac'])
                c_customs = find_c(['海关号', 'hải quan', 'hai quan'])
                c_fab_code = find_c(['布编号'])
                c_fab_type = find_c(['布种', 'loại vải', 'loai vai'])
                c_comp = find_c(['成分', 'thành phần', 'thanh phan'])
                c_weight = find_c(['重量', 'kg'])
                c_rolls = find_c(['匹数', 'cay', 'cây'])
                c_yards = find_c(['码数', 'yard'])
                c_invoice = find_c(['发票号', 'invoice'])
                c_notes = find_c(['备注', 'ghi chú', 'ghi chu'])

                conn = get_db()
                cur = conn.cursor()
                inserted_cnt = 0
                for _, r in df_read.iterrows():
                    l_val = str(r[c_lot]).strip() if c_lot and pd.notna(r[c_lot]) else ""
                    s_val = str(r[c_style]).strip() if c_style and pd.notna(r[c_style]) else ""
                    cu_val = str(r[c_cust]).strip() if c_cust and pd.notna(r[c_cust]) else ""
                    co_val = str(r[c_color]).strip() if c_color and pd.notna(r[c_color]) else ""

                    if not l_val and not s_val:
                        continue

                    if cu_val:
                        cur.execute("INSERT OR IGNORE INTO customers (name) VALUES (?)", (cu_val,))

                    d_val = format_excel_date(r[c_date] if c_date and pd.notna(r[c_date]) else "")
                    cust_no = str(r[c_customs]) if c_customs and pd.notna(r[c_customs]) else ""
                    code_no = str(r[c_fab_code]) if c_fab_code and pd.notna(r[c_fab_code]) else ""
                    type_no = str(r[c_fab_type]) if c_fab_type and pd.notna(r[c_fab_type]) else ""
                    comp_no = str(r[c_comp]) if c_comp and pd.notna(r[c_comp]) else ""

                    try: w_no = float(r[c_weight]) if c_weight and pd.notna(r[c_weight]) else 0.0
                    except Exception: w_no = 0.0

                    try: ro_no = float(r[c_rolls]) if c_rolls and pd.notna(r[c_rolls]) else 0.0
                    except Exception: ro_no = 0.0

                    try: yd_no = float(r[c_yards]) if c_yards and pd.notna(r[c_yards]) else 0.0
                    except Exception: yd_no = 0.0

                    inv_no = str(r[c_invoice]) if c_invoice and pd.notna(r[c_invoice]) else ""
                    nt_no = str(r[c_notes]) if c_notes and pd.notna(r[c_notes]) else ""

                    cur.execute('''
                        INSERT INTO inventory_stock (
                            customer, in_date, style, lot, color, customs_no, fabric_code,
                            fabric_type, composition, weight, rolls, yards, invoice, notes
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (cu_val, d_val, s_val, l_val, co_val, cust_no, code_no, type_no, comp_no, w_no, ro_no, yd_no, inv_no, nt_no))
                    inserted_cnt += 1

                conn.commit()
                conn.close()
                st.success(f"✓ Đã nạp thành công {inserted_cnt} dòng dữ liệu nhập kho từ Excel vào hệ thống!")
                st.rerun()
            except Exception as e:
                st.error(f"Lỗi khi xử lý file Excel: {e}")

        selected_suggest_val = st.selectbox(
            "LOT / 缸号 * (Tự đề xuất từ Dữ liệu Nhập Kho):",
            options=suggest_labels,
            index=0
        )

        preset_lot = ""
        preset_style = ""
        preset_color = ""
        if selected_suggest_val in suggest_map:
            sel_info = suggest_map[selected_suggest_val]
            preset_lot = sel_info["lot"]
            preset_style = sel_info["style"]
            preset_color = sel_info["color"]
            if sel_info["customer"]:
                st.session_state.selected_customer = sel_info["customer"]

        col_lot, col_style = st.columns(2)
        with col_lot:
            lot_input = st.text_input("LOT / 缸号 *:", value=preset_lot, placeholder="Gõ LOT để chọn từ dữ liệu nhập kho...")
        with col_style:
            style_input = st.text_input("MÃ HÀNG / 款号 *:", value=preset_style, placeholder="Mã hàng...")

        color_input = st.text_input("MÀU SẮC / 颜色:", value=preset_color, placeholder="Màu sắc...")

        cust_list = get_customers_list()
        current_cust = st.session_state.selected_customer
        if current_cust not in cust_list and cust_list:
            current_cust = cust_list[0]
            st.session_state.selected_customer = current_cust

        c_lbl_left, c_lbl_right = st.columns([6, 6])
        with c_lbl_left:
            st.markdown("<b style='font-size:13px; color:#0f172a;'>KHÁCH HÀNG / 客户:</b>", unsafe_allow_html=True)
        with c_lbl_right:
            st.markdown(f"<div style='text-align:right; font-weight:700; color:#4338ca;'>Đã chọn: {current_cust}</div>", unsafe_allow_html=True)

        grid_cols = st.columns(3)
        for i, c_name in enumerate(cust_list):
            is_active = (c_name == current_cust)
            btn_type = "primary" if is_active else "secondary"
            with grid_cols[i % 3]:
                if st.button(c_name, key=f"btn_cust_{c_name}", type=btn_type, use_container_width=True):
                    st.session_state.selected_customer = c_name
                    st.rerun()

        act_c1, act_c2 = st.columns(2)
        with act_c1:
            with st.expander("+ Thêm KH Mới / 添加客户"):
                new_c_val = st.text_input("Tên KH mới:", key="in_new_cust")
                if st.button("Lưu KH mới"):
                    if new_c_val.strip():
                        conn = get_db()
                        cur = conn.cursor()
                        cur.execute("INSERT OR IGNORE INTO customers (name) VALUES (?)", (new_c_val.strip(),))
                        conn.commit()
                        conn.close()
                        st.session_state.selected_customer = new_c_val.strip()
                        st.success(f"Đã thêm: {new_c_val.strip()}")
                        st.rerun()
        with act_c2:
            with st.expander("- Xoá KH / 删除客户"):
                del_c_val = st.selectbox("Chọn KH cần xoá:", options=cust_list, key="sb_del_cust")
                if st.button("Xác nhận xoá KH"):
                    conn = get_db()
                    cur = conn.cursor()
                    cur.execute("DELETE FROM customers WHERE name=?", (del_c_val,))
                    conn.commit()
                    conn.close()
                    st.session_state.selected_customer = ""
                    st.warning(f"Đã xoá khách hàng: {del_c_val}")
                    st.rerun()

        current_grade = st.session_state.selected_grade
        g_lbl_left, g_lbl_right = st.columns([6, 6])
        with g_lbl_left:
            st.markdown("<b style='font-size:13px; color:#0f172a;'>XẾP LOẠI / 布料等级:</b>", unsafe_allow_html=True)
        with g_lbl_right:
            st.markdown(f"<div style='text-align:right; font-weight:700; color:#059669;'>Đã chọn: {current_grade}</div>", unsafe_allow_html=True)

        grade_cols = st.columns(4)
        grades_arr = ["Loại A (A 级)", "Loại B (B 级)", "Loại C (C 级)", "Loại D (D 级)"]
        for g_i, g_val in enumerate(grades_arr):
            is_active_g = (g_val == current_grade)
            g_btn_type = "primary" if is_active_g else "secondary"
            with grade_cols[g_i]:
                if st.button(g_val, key=f"btn_grade_{g_val}", type=g_btn_type, use_container_width=True):
                    st.session_state.selected_grade = g_val
                    st.rerun()

        notes_input = st.text_area("Ghi chú / 检验备注:", height=60, placeholder="Ghi chú kết quả kiểm...")

        today_code = datetime.now().strftime("%Y%m%d")
        safe_c = sanitize_for_filename(current_cust) or "KhongKH"
        safe_l = sanitize_for_filename(lot_input) or "lot"
        safe_s = sanitize_for_filename(style_input) or "ma hang"
        safe_col = sanitize_for_filename(color_input) or "mau sac"
        preview_filename = f"{today_code} - {safe_c} - {safe_l} - {safe_s} - {safe_col} - anh goc.jpg"

        st.markdown("<div style='font-size:11px; color:#64748b;'>Quy chuẩn tên file lưu (Tiếng Việt không dấu):</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='filename-box'>{preview_filename}</div>", unsafe_allow_html=True)

        # LẤY ĐÚNG THƯ MỤC CẤU HÌNH ĐỂ LƯU THỰC TẾ
        target_storage_dir = get_storage_dir_setting()
        try:
            os.makedirs(target_storage_dir, exist_ok=True)
        except Exception:
            target_storage_dir = "storage"
            os.makedirs(target_storage_dir, exist_ok=True)

        if st.button("💾 XÁC NHẬN _LƯU BÁO CÁO / 确认并保存", type="primary", use_container_width=True):
            if not lot_input.strip() or not style_input.strip():
                st.error("Vui lòng nhập LOT và MÃ HÀNG!")
            elif st.session_state.main_image_cache is None:
                st.error("Vui lòng chọn ảnh báo cáo chính trước khi lưu!")
            else:
                main_file_path = os.path.join(target_storage_dir, preview_filename)
                st.session_state.main_image_cache.save(main_file_path, "JPEG", quality=95)

                defect_saved_paths = []
                if uploaded_defects:
                    for d_idx, d_file in enumerate(uploaded_defects, 1):
                        d_name = f"{today_code} - {safe_c} - {safe_l} - {safe_s} - {safe_col} - loi {d_idx}.jpg"
                        d_path = os.path.join(target_storage_dir, d_name)
                        with open(d_path, "wb") as f_out:
                            f_out.write(d_file.getbuffer())
                        defect_saved_paths.append(d_path)

                conn = get_db()
                cur = conn.cursor()
                cur.execute('''
                    INSERT INTO reports (lot, style, color, customer, grade, notes, main_image, defect_images, defect_count, upload_date)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    lot_input.strip(), style_input.strip(), color_input.strip(), current_cust,
                    current_grade, notes_input.strip(), main_file_path, json.dumps(defect_saved_paths),
                    len(defect_saved_paths), today_code
                ))
                conn.commit()
                conn.close()

                st.success(f"🎉 Đã lưu báo cáo thành công vào: {main_file_path}")
                st.session_state.main_image_cache = None
                st.rerun()

        st.markdown(f"<div style='font-size:12px; color:#b45309; margin-top:8px;'>📁 Thư mục lưu: <b>{target_storage_dir}</b></div>", unsafe_allow_html=True)

    if st.session_state.get("open_inventory_modal", False):
        st.divider()
        st.subheader("📋 DANH SÁCH NHẬP KHO HIỆN TẠI / 现存入库清单")
        conn = get_db()
        df_inv_full = pd.read_sql_query("SELECT id, customer, in_date, lot, style, color, customs_no, fabric_type, rolls, weight, yards FROM inventory_stock ORDER BY id DESC", conn)
        conn.close()
        st.dataframe(df_inv_full, use_container_width=True, height=350)
        if st.button("Đóng danh sách nhập kho"):
            st.session_state.open_inventory_modal = False
            st.rerun()

# ==============================================================================
# TAB 2: TỔNG QUAN
# ==============================================================================
elif nav_option == "📊 Tổng quan / 概览":
    st.header("📊 TỔNG QUAN BÁO CÁO KIỂM VẢI / 验布报告概览", divider="blue")

    conn = get_db()
    df_all_reports = pd.read_sql_query("SELECT * FROM reports ORDER BY id DESC", conn)
    conn.close()

    total_cnt = len(df_all_reports)
    today_code = datetime.now().strftime("%Y%m%d")
    today_cnt = len(df_all_reports[df_all_reports["upload_date"] == today_code]) if total_cnt > 0 else 0
    unique_lots_cnt = df_all_reports["lot"].nunique() if total_cnt > 0 else 0

    c1, c2, c3 = st.columns(3)
    c1.metric("Tổng báo cáo / 报告总数", f"{total_cnt}", delta="Báo cáo")
    c2.metric("Hôm nay / 今日新增", f"{today_cnt}", delta="Hôm nay")
    c3.metric("LOT đã quản lý / 已管缸号", f"{unique_lots_cnt}", delta="LOT")

    st.subheader("THỐNG KÊ THEO KHÁCH HÀNG / 按客户统计:")
    custs = get_customers_list()
    if custs:
        grid_c = st.columns(5)
        for i, c_name in enumerate(custs):
            cnt = len(df_all_reports[df_all_reports["customer"] == c_name]) if total_cnt > 0 else 0
            pct = round((cnt / total_cnt * 100), 1) if total_cnt > 0 else 0.0
            border_col = "#6366f1" if cnt > 0 else "#cbd5e1"
            bg_col = "#eef2ff" if cnt > 0 else "#ffffff"

            with grid_c[i % 5]:
                st.markdown(f"""
                    <div style="background:{bg_col}; border:1px solid {border_col}; border-radius:6px; padding:8px 10px; margin-bottom:8px;">
                        <b style="color:#0f172a; font-size:13px;">{c_name}</b><br>
                        <span style="color:#64748b; font-size:11px;">{cnt} báo cáo ({pct}%)</span>
                    </div>
                """, unsafe_allow_html=True)

    st.subheader("Báo cáo kiểm vải gần đây / 最近报告:")
    if not df_all_reports.empty:
        display_df = df_all_reports[["upload_date", "lot", "style", "color", "customer", "grade", "defect_count"]].head(10)
        display_df.columns = ["Ngày / 日期", "LOT (缸号)", "MÃ HÀNG (款号)", "MÀU SẮC (颜色)", "KHÁCH HÀNG (客户)", "XẾP LOẠI (等级)", "Số Lỗi"]
        st.dataframe(display_df, use_container_width=True)
    else:
        st.info("Chưa có báo cáo nào trong hệ thống.")

# ==============================================================================
# TAB 3: TRA CỨU
# ==============================================================================
elif nav_option == "🔍 Tra cứu / 查询":
    st.header("🔍 TRA CỨU & ĐÓNG GÓI BÁO CÁO / 报告查询与打包", divider="violet")

    conn = get_db()
    df_rep = pd.read_sql_query("SELECT id, lot, style, color, customer, grade, defect_count, upload_date, main_image, defect_images, notes FROM reports ORDER BY id DESC", conn)
    conn.close()

    fl1, fl2, fl3 = st.columns(3)
    with fl1:
        kw_lot = st.text_input("Tìm kiếm LOT:", "")
    with fl2:
        kw_style = st.text_input("Mã hàng:", "")
    with fl3:
        all_c_options = ["-- Tất cả khách hàng --"] + get_customers_list()
        kw_cust = st.selectbox("Khách hàng:", options=all_c_options)

    filtered_df = df_rep
    if kw_lot.strip():
        filtered_df = filtered_df[filtered_df["lot"].str.lower().str.contains(kw_lot.strip().lower(), na=False)]
    if kw_style.strip():
        filtered_df = filtered_df[filtered_df["style"].str.lower().str.contains(kw_style.strip().lower(), na=False)]
    if kw_cust != "-- Tất cả khách hàng --":
        filtered_df = filtered_df[filtered_df["customer"] == kw_cust]

    st.write(f"Tìm thấy: **{len(filtered_df)}** báo cáo")

    if not filtered_df.empty:
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for _, r_item in filtered_df.iterrows():
                m_img_p = r_item["main_image"]
                if m_img_p and os.path.exists(m_img_p):
                    zf.write(m_img_p, arcname=os.path.basename(m_img_p))
                if r_item["defect_images"]:
                    try:
                        for d_img_p in json.loads(r_item["defect_images"]):
                            if os.path.exists(d_img_p):
                                zf.write(d_img_p, arcname=os.path.basename(d_img_p))
                    except Exception:
                        pass
        zip_buf.seek(0)

        st.download_button(
            label="📦 ĐÓNG GÓI BÁO CÁO ĐANG LỌC (.ZIP) / 打包下载",
            data=zip_buf,
            file_name=f"BaoCaoKiemVai_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
            mime="application/zip",
            type="primary"
        )

    st.divider()
    target_storage_dir = get_storage_dir_setting()
    for _, rep in filtered_df.iterrows():
        with st.container():
            c_info, c_action = st.columns([7, 3])
            with c_info:
                st.markdown(f"#### LOT: `{rep['lot']}` | Mã: `{rep['style']}` | KH: **{rep['customer']}** | {rep['grade']}")
                st.caption(f"Ngày kiểm: {rep['upload_date']} | Màu: {rep['color']} | Số ảnh lỗi: {rep['defect_count']} | Ghi chú: {rep['notes'] or 'Không có'}")

                if rep["main_image"] and os.path.exists(rep["main_image"]):
                    with st.expander("👁️ Xem chi tiết ảnh báo cáo và các ảnh lỗi"):
                        st.image(rep["main_image"], caption=f"Ảnh gốc: {os.path.basename(rep['main_image'])}", use_column_width=True)
                        if rep["defect_images"]:
                            try:
                                d_arr = json.loads(rep["defect_images"])
                                if d_arr:
                                    st.write(f"**Danh sách {len(d_arr)} ảnh lỗi:**")
                                    img_cols = st.columns(min(4, len(d_arr)))
                                    for idx_d, d_file in enumerate(d_arr):
                                        if os.path.exists(d_file):
                                            with img_cols[idx_d % 4]:
                                                st.image(d_file, caption=f"Lỗi {idx_d+1}", use_column_width=True)
                            except Exception:
                                pass
            with c_action:
                with st.expander("📷 Bổ sung / Sửa / Xoá"):
                    new_main_append = st.file_uploader("Đổi ảnh gốc:", type=["jpg", "png"], key=f"append_main_{rep['id']}")
                    new_def_append = st.file_uploader("Thêm ảnh lỗi:", type=["jpg", "png"], accept_multiple_files=True, key=f"append_def_{rep['id']}")

                    if st.button("Lưu ảnh bổ sung", key=f"btn_save_append_{rep['id']}"):
                        t_code = datetime.now().strftime("%Y%m%d")
                        s_c = sanitize_for_filename(rep['customer']) or "KhongKH"
                        s_l = sanitize_for_filename(rep['lot'])
                        s_s = sanitize_for_filename(rep['style'])
                        s_col = sanitize_for_filename(rep['color']) or "KhongMau"

                        cur_main = rep["main_image"]
                        if new_main_append:
                            cur_main = os.path.join(target_storage_dir, f"{t_code} - {s_c} - {s_l} - {s_s} - {s_col} - anh goc.jpg")
                            with open(cur_main, "wb") as f_out:
                                f_out.write(new_main_append.getbuffer())

                        cur_defects = []
                        if rep["defect_images"]:
                            try: cur_defects = json.loads(rep["defect_images"])
                            except Exception: pass

                        if new_def_append:
                            st_idx = len(cur_defects) + 1
                            for n_i, f in enumerate(new_def_append):
                                d_path = os.path.join(target_storage_dir, f"{t_code} - {s_c} - {s_l} - {s_s} - {s_col} - loi {st_idx + n_i}.jpg")
                                with open(d_path, "wb") as f_out:
                                    f_out.write(f.getbuffer())
                                cur_defects.append(d_path)

                        conn = get_db()
                        cur = conn.cursor()
                        cur.execute('''
                            UPDATE reports SET main_image=?, defect_images=?, defect_count=? WHERE id=?
                        ''', (cur_main, json.dumps(cur_defects), len(cur_defects), rep['id']))
                        conn.commit()
                        conn.close()

                        st.success("Đã bổ sung ảnh thành công!")
                        st.rerun()

                    if st.button("🗑️ Xoá báo cáo", key=f"btn_del_rep_{rep['id']}", type="secondary"):
                        conn = get_db()
                        cur = conn.cursor()
                        cur.execute("DELETE FROM reports WHERE id=?", (rep['id'],))
                        conn.commit()
                        conn.close()
                        st.warning("Đã xoá báo cáo!")
                        st.rerun()
            st.divider()

# ==============================================================================
# TAB 4: DANH SÁCH NHẬP KHO
# ==============================================================================
elif nav_option == "📋 Danh Sách Nhập Kho / 入库清单":
    st.header("📋 DANH SÁCH NHẬP KHO / 入库清单管理", divider="rainbow")

    conn = get_db()
    df_stock = pd.read_sql_query("SELECT id, customer, in_date, lot, style, color, customs_no, fabric_type, rolls, weight, yards, invoice FROM inventory_stock ORDER BY id DESC", conn)
    conn.close()

    st.write(f"Tổng số dòng hiện có: **{len(df_stock)}** dòng")
    kw_inv_search = st.text_input("🔍 Tìm kiếm theo LOT, Mã hàng hoặc Khách hàng:", "")

    if kw_inv_search.strip():
        kw_s = kw_inv_search.strip().lower()
        mask = (
            df_stock["lot"].str.lower().str.contains(kw_s, na=False) |
            df_stock["style"].str.lower().str.contains(kw_s, na=False) |
            df_stock["customer"].str.lower().str.contains(kw_s, na=False) |
            df_stock["color"].str.lower().str.contains(kw_s, na=False)
        )
        display_stock = df_stock[mask]
    else:
        display_stock = df_stock

    st.dataframe(display_stock, use_container_width=True, height=450)

    with st.expander("🛠️ Xoá dữ liệu nhập kho"):
        c_del1, c_del2 = st.columns([4, 6])
        with c_del1:
            id_to_del = st.number_input("Nhập ID cần xoá:", min_value=1, step=1)
            if st.button("Xoá ID này"):
                conn = get_db()
                cur = conn.cursor()
                cur.execute("DELETE FROM inventory_stock WHERE id=?", (id_to_del,))
                conn.commit()
                conn.close()
                st.success(f"Đã xoá dòng ID: {id_to_del}")
                st.rerun()
        with c_del2:
            if st.button("⚠️ Xoá TOÀN BỘ danh mục nhập kho"):
                conn = get_db()
                cur = conn.cursor()
                cur.execute("DELETE FROM inventory_stock")
                conn.commit()
                conn.close()
                st.warning("Đã xoá toàn bộ dữ liệu nhập kho.")
                st.rerun()

# ==============================================================================
# TAB 5: CÀI ĐẶT & BACKUP (ĐÃ SỬA LỖI CHỌN ĐƯỜNG DẪN THƯ MỤC)
# ==============================================================================
elif nav_option == "⚙ Cài đặt _Thư mục / 设置":
    st.header("⚙ CÀI ĐẶT & SAO LƯU DỰ PHÒNG / 设置与备份", divider="orange")

    st.subheader("1. Cài đặt đường dẫn thư mục lưu / 存储目录配置")
    current_dir_cfg = get_storage_dir_setting()

    # Khung chọn nhanh bằng hộp thoại hoặc nhập tay
    c_set_in, c_set_dlg = st.columns([7, 3])
    with c_set_in:
        new_dir_input = st.text_input(
            "Đường dẫn lưu báo cáo (Có thể gõ trực tiếp ổ D:/, E:/...):",
            value=current_dir_cfg,
            key="input_storage_dir_box"
        )
    with c_set_dlg:
        st.write("&nbsp;", unsafe_allow_html=True)
        if st.button("📁 Mở hộp thoại chọn thư mục", use_container_width=True):
            picked_path = open_native_folder_dialog(current_dir_cfg)
            if picked_path:
                set_storage_dir_setting(picked_path)
                st.success(f"✓ Đã chọn thư mục: {picked_path}")
                st.rerun()

    # Nút lưu cấu hình và kiểm tra tạo thư mục
    c_save_dir, c_check_status = st.columns([3, 7])
    with c_save_dir:
        if st.button("💾 Lưu cấu hình thư mục", type="primary", use_container_width=True):
            cleaned_p = set_storage_dir_setting(new_dir_input)
            st.success(f"✓ Đã cập nhật và tạo thư mục thành công: `{cleaned_p}`")
            st.rerun()

    with c_check_status:
        if os.path.exists(current_dir_cfg):
            st.markdown(f"<span style='color:#059669; font-weight:600;'>✓ Trạng thái: Thư mục đã sẵn sàng tồn tại trên ổ đĩa (`{current_dir_cfg}`)</span>", unsafe_allow_html=True)
        else:
            st.markdown(f"<span style='color:#dc2626; font-weight:600;'>⚠️ Trạng thái: Thư mục chưa tồn tại (Bấm 'Lưu cấu hình thư mục' để hệ thống tự tạo)</span>", unsafe_allow_html=True)

    # Thư mục gợi ý phổ biến
    st.caption("💡 Gợi ý đường dẫn nhanh trên Windows: `D:/BaoCaoKiemVai/Original`, `E:/HinhAnhBaoCaoKiemVai`, `D:/FabricInspection`")

    st.divider()

    # Sao lưu & Phục hồi JSON
    st.subheader("2. Sao lưu & Phục hồi cơ sở dữ liệu / 数据库备份与恢复")
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM reports")
    all_reps = cur.fetchall()
    cur.execute("SELECT * FROM inventory_stock")
    all_stocks = cur.fetchall()
    cur.execute("SELECT name FROM customers")
    all_c_names = [r[0] for r in cur.fetchall()]
    conn.close()

    backup_bundle = {
        "version": "Streamlit-2.0",
        "exported_at": datetime.now().isoformat(),
        "storage_dir": current_dir_cfg,
        "customers": all_c_names,
        "reports": all_reps,
        "inventory": all_stocks
    }
    json_export_bytes = json.dumps(backup_bundle, ensure_ascii=False, indent=2).encode('utf-8')

    st.download_button(
        label="💾 XUẤT FILE SAO LƯU (.JSON) / 导出备份",
        data=json_export_bytes,
        file_name=f"Backup_BaoCaoKiemVai_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
        mime="application/json",
        type="primary"
    )

    restore_upload = st.file_uploader("📥 Khôi phục dữ liệu từ file .JSON:", type=["json"])
    if restore_upload:
        try:
            pkg = json.load(restore_upload)
            st.info(f"Tìm thấy: {len(pkg.get('reports', []))} báo cáo và {len(pkg.get('inventory', []))} dòng nhập kho.")
            if st.button("Xác nhận nạp đè dữ liệu sao lưu"):
                conn = get_db()
                cur = conn.cursor()
                for c in pkg.get("customers", []):
                    cur.execute("INSERT OR IGNORE INTO customers (name) VALUES (?)", (c,))
                for r in pkg.get("reports", []):
                    cur.execute('''
                        INSERT INTO reports (lot, style, color, customer, grade, notes, main_image, defect_images, defect_count, upload_date)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10]))
                for inv in pkg.get("inventory", []):
                    cur.execute('''
                        INSERT INTO inventory_stock (customer, in_date, style, lot, color, customs_no, fabric_code, fabric_type, composition, weight, rolls, yards, invoice, notes)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (inv[1], inv[2], inv[3], inv[4], inv[5], inv[6], inv[7], inv[8], inv[9], inv[10], inv[11], inv[12], inv[13], inv[14]))
                conn.commit()
                conn.close()
                st.success("Đã khôi phục toàn bộ dữ liệu thành công!")
                st.rerun()
        except Exception as e:
            st.error(f"Lỗi file sao lưu: {e}")