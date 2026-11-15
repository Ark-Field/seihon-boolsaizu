import streamlit as st
import pandas as pd
import requests
import io
import os
import math
import urllib.request
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.graphics.shapes import Drawing, Rect, String, Line
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import openpyxl

# ページ基本設定
st.set_page_config(page_title="芯ボール・背ボール・溝あき対応 発注書・図面自動作成システム", layout="wide")

# カスタムスタイル（ダークネイビー調）
st.markdown("""
<style>
    .stButton>button {
        background-color: #1e3d59;
        color: white;
        border-radius: 4px;
    }
    .stButton>button:hover {
        background-color: #172b3a;
        color: white;
    }
</style>
""", unsafe_allow_html=True)

st.title("📦 芯ボール・背ボール・溝あき対応 発注書作成")
st.markdown("💡 DBpocket（）：サイドバーから受注番号を指定してデータを取得してください。")

def safe_int(val, default=0):
    if pd.isna(val) or str(val).strip() == '' or str(val).lower() == 'nan':
        return default
    try:
        return int(float(str(val)))
    except (ValueError, TypeError):
        return default

def safe_float(val, default=0.0):
    if pd.isna(val) or str(val).strip() == '' or str(val).lower() == 'nan':
        return default
    try:
        return float(str(val))
    except (ValueError, TypeError):
        return default

# フォント設定
font_name = 'Helvetica'
font_filename = "NotoSansJP-Regular.ttf"

if not os.path.exists(font_filename):
    try:
        font_url = "https://github.com/google/fonts/raw/main/ofl/notosansjp/NotoSansJP-Regular.ttf"
        urllib.request.urlretrieve(font_url, font_filename)
    except Exception as e:
        print(f"Font download error: {e}")

if os.path.exists(font_filename):
    try:
        pdfmetrics.registerFont(TTFont('JapFont', font_filename))
        font_name = 'JapFont'
    except Exception as e:
        print(f"Font registration error: {e}")

# --- URLパラメータの取得 ---
query_params = st.query_params
url_order_id = query_params.get("order_id", "")

# --- Session State の初期化 ---
if "loaded_data" not in st.session_state:
    st.session_state.loaded_data = None
if "app29_data" not in st.session_state:
    st.session_state.app29_data = None
if "parent_order_id" not in st.session_state:
    st.session_state.parent_order_id = ""
if "order_id" not in st.session_state:
    st.session_state.order_id = url_order_id
if "book_title" not in st.session_state:
    st.session_state.book_title = ""
if "quantity" not in st.session_state:
    st.session_state.quantity = 0
if "supplier_cover" not in st.session_state:
    st.session_state.supplier_cover = ""
if "due_cover" not in st.session_state:
    st.session_state.due_cover = ""
if "ampm_cover" not in st.session_state:
    st.session_state.ampm_cover = ""
if "supplier_foil" not in st.session_state:
    st.session_state.supplier_foil = ""
if "due_foil" not in st.session_state:
    st.session_state.due_foil = ""
if "ampm_foil" not in st.session_state:
    st.session_state.ampm_foil = ""
if "delivery_to" not in st.session_state:
    st.session_state.delivery_to = ""
if "hon_h" not in st.session_state:
    st.session_state.hon_h = 0.0
if "hon_w" not in st.session_state:
    st.session_state.hon_w = 0.0
if "zenpan_h" not in st.session_state:
    st.session_state.zenpan_h = 1100.0
if "zenpan_w" not in st.session_state:
    st.session_state.zenpan_w = 800.0
if "parts_rows" not in st.session_state:
    st.session_state.parts_rows = []
if "app29_parts_rows" not in st.session_state:
    st.session_state.app29_parts_rows = []

def get_field_val(rec, field_id, default=""):
    if not rec or field_id not in rec:
        return default
    val = rec[field_id]
    if isinstance(val, dict):
        return val.get("value", default)
    return val if val is not None else default

# --- APIデータ取得＆連携関数 ---
def fetch_and_store_pocket_data(order_num):
    try:
        # 1. アプリ50からデータを取得
        api_key_50 = st.secrets["app50"]["api_key"]
        url_50 = "https://app050.at-pocket.com/seihon02_bb/api/apps/50/records"
        headers_50 = {"X-At-Pocket-API-Key": api_key_50, "Content-Type": "application/json"}
        
        response_50 = requests.get(url_50, headers=headers_50, timeout=30)
        if response_50.status_code == 200:
            records_50 = response_50.json().get("records", [])
            clean_target = str(order_num).replace("JS", "").strip()
            
            target_rec_50 = None
            for rec in records_50:
                inner_rec = rec.get("record", rec)
                f55 = str(get_field_val(inner_rec, "field-55")).strip()
                if clean_target in [f55, f55.replace("JS", "")]:
                    target_rec_50 = rec
                    break
            
            if not target_rec_50:
                for rec in records_50:
                    inner_rec = rec.get("record", rec)
                    for k, v in inner_rec.items():
                        v_str = str(v.get("value", v) if isinstance(v, dict) else v)
                        if clean_target and clean_target in v_str.replace("JS", "").strip():
                            target_rec_50 = rec
                            break
                    if target_rec_50:
                        break
            
            if target_rec_50:
                st.session_state.loaded_data = target_rec_50
                r50 = target_rec_50.get("record", target_rec_50)
                
                st.session_state.order_id = str(get_field_val(r50, "field-55", order_num))
                parent_id = str(get_field_val(r50, "field-23", "")).strip()
                st.session_state.parent_order_id = parent_id
                
                st.session_state.book_title = str(get_field_val(r50, "field-24", ""))
                st.session_state.quantity = safe_int(get_field_val(r50, "field-70", get_field_val(r50, "field-25", 0)), 0)
                
                st.session_state.hon_h = safe_float(get_field_val(r50, "field-27", 0.0), 0.0)
                st.session_state.hon_w = safe_float(get_field_val(r50, "field-28", 0.0), 0.0)
                st.session_state.zenpan_h = safe_float(get_field_val(r50, "field-67", 1100.0), 1100.0)
                st.session_state.zenpan_w = safe_float(get_field_val(r50, "field-68", 800.0), 800.0)
                
                st.session_state.delivery_to = str(get_field_val(r50, "field-69", ""))
                st.session_state.supplier_cover = str(get_field_val(r50, "field-57", ""))
                st.session_state.due_cover = str(get_field_val(r50, "field-61", ""))
                st.session_state.ampm_cover = str(get_field_val(r50, "field-62", ""))
                st.session_state.supplier_foil = str(get_field_val(r50, "field-59", ""))
                st.session_state.due_foil = str(get_field_val(r50, "field-60", ""))
                st.session_state.ampm_foil = str(get_field_val(r50, "field-63", ""))

                # 2. 【アプリ29（親データ）】の取得処理（クエリピンポイント検索）
                app29_rows = []
                if parent_id:
                    try:
                        api_key_29 = st.secrets["app29"]["api_key"]
                        url_29 = "https://app050.at-pocket.com/seihon02_bb/api/apps/29/records"
                        headers_29 = {"X-At-Pocket-API-Key": api_key_29, "Content-Type": "application/json"}
                        
                        clean_parent = parent_id.replace("JS", "").strip()
                        query_str = f'field-2 = "{clean_parent}"'
                        response_29 = requests.get(url_29, headers=headers_29, params={"query": query_str}, timeout=30)
                        
                        target_rec_29 = None
                        if response_29.status_code == 200:
                            records_29 = response_29.json().get("records", [])
                            if records_29:
                                target_rec_29 = records_29[0]
                        
                        if not target_rec_29:
                            response_all = requests.get(url_29, headers=headers_29, timeout=30)
                            if response_all.status_code == 200:
                                for prec in response_all.json().get("records", []):
                                    prec_inner = prec.get("record", prec)
                                    for pk, pv in prec_inner.items():
                                        pv_str = str(pv.get("value", pv) if isinstance(pv, dict) else pv)
                                        if clean_parent and clean_parent in pv_str.replace("JS", "").strip():
                                            target_rec_29 = prec
                                            break
                                    if target_rec_29:
                                        break

                        if target_rec_29:
                            st.session_state.app29_data = target_rec_29
                            r29 = target_rec_29.get("record", target_rec_29)
                            
                            f31_raw = r29.get("field-31", [])
                            if isinstance(f31_raw, dict):
                                f31_raw = f31_raw.get("value", [])
                            
                            if isinstance(f31_raw, list):
                                for sub_row in f31_raw:
                                    sub_rec = sub_row.get("record", sub_row) if isinstance(sub_row, dict) else sub_row
                                    if not isinstance(sub_rec, dict):
                                        continue
                                    def get_sub29_val(key):
                                        val = sub_rec.get(key, "")
                                        if isinstance(val, dict):
                                            return val.get("value", "")
                                        return val if val is not None else ""
                                    
                                    app29_rows.append({
                                        "PDF掲載": True,  # デフォルトでチェックON
                                        "種目": str(get_sub29_val("field-31_6")).strip(),
                                        "内容": str(get_sub29_val("field-31_1")).strip()
                                    })
                    except Exception as e2:
                        print(f"アプリ29取得エラー: {e2}")

                st.session_state.app29_parts_rows = app29_rows

                # 3. アプリ50の計算用パーツリスト（field-64）の取得
                app50_parts = []
                f64_raw = r50.get("field-64", [])
                if isinstance(f64_raw, dict):
                    f64_raw = f64_raw.get("value", [])
                
                if isinstance(f64_raw, list):
                    for sub_row in f64_raw:
                        sub_rec = sub_row.get("record", sub_row) if isinstance(sub_row, dict) else sub_row
                        if not isinstance(sub_rec, dict):
                            continue
                        def get_sub_val(key):
                            val = sub_rec.get(key, "")
                            if isinstance(val, dict):
                                return val.get("value", "")
                            return val if val is not None else ""
                        b_name = str(get_sub_val("field-64_2")).strip()
                        if not b_name:
                            continue
                        is_groove = "溝あき" in b_name
                        app50_parts.append({
                            "部材名": b_name,
                            "銘柄": str(get_sub_val("field-64_1")).strip(),
                            "備考": "",
                            "１冊面数": 0 if is_groove else safe_int(get_sub_val("field-64_4"), 1),
                            "天地サイズ+-": safe_float(get_sub_val("field-64_5"), 0.0),
                            "左右サイズ+-": safe_float(get_sub_val("field-64_6"), 0.0),
                            "固定幅【背】": safe_float(get_sub_val("field-64_7"), 0.0),
                            "レコードID": safe_int(sub_row.get("recordSubId", 1), 1)
                        })

                st.session_state.parts_rows = app50_parts
                return True
    except Exception as e:
        print(f"API取得エラー: {e}")
    return False

# --- サイドバー：手動検索機能 ---
with st.sidebar:
    st.header("🔍 DBpocket（アプリ50）検索")
    search_input = st.text_input("受注番号を入力（例: 260543）", value=st.session_state.order_id)
    search_button = st.button("APIからデータを取得する")

if search_button and search_input:
    with st.spinner(f"⏳ 受注番号「{search_input}」およびアプリ29親データを取得中..."):
        success = fetch_and_store_pocket_data(search_input)
        if success:
            st.sidebar.success(f"🔗 データロード成功！（親ID: {st.session_state.parent_order_id}）")
            st.rerun()
        else:
            st.sidebar.warning(f"⚠️ アプリ50内に「{search_input}」が見つかりませんでした。")

# 自動URLパラメータ連携
if url_order_id and not st.session_state.loaded_data:
    fetch_and_store_pocket_data(url_order_id)

# --- 画面レイアウト ---
if not st.session_state.loaded_data:
    st.info("ℹ️ サイドバーに受注番号を入力して **[APIからデータを取得する]** ボタンを押してください。")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. 受注・発注基本情報")
    order_id = st.text_input("受注番号", value=st.session_state.order_id)
    book_title = st.text_input("書名（製品名）", value=st.session_state.book_title)
    quantity = st.number_input("必要数量（受注部数）", min_value=0, value=st.session_state.quantity)
    supplier_cover = st.text_input("表紙貼外注先", value=st.session_state.supplier_cover)
    due_cover = st.text_input("納期（表紙）", value=st.session_state.due_cover)
    ampm_cover = st.text_input("AMPM（表紙）", value=st.session_state.ampm_cover)
    supplier_foil = st.text_input("表紙箔押外注先", value=st.session_state.supplier_foil)
    due_foil = st.text_input("納期（箔押）", value=st.session_state.due_foil)
    ampm_foil = st.text_input("AMPM（箔押）", value=st.session_state.ampm_foil)
    delivery_to = st.text_input("納品先", value=st.session_state.delivery_to)

with col2:
    st.subheader("2. 本文仕様 ＆ 全判材料サイズ設定")
    hon_h = st.number_input("本文 天地サイズ (mm)", min_value=0.0, value=st.session_state.hon_h)
    hon_w = st.number_input("本文 左右サイズ (mm)", min_value=0.0, value=st.session_state.hon_w)
    zenpan_h = st.number_input("材料規格 全判天地 (mm)", min_value=0.0, value=st.session_state.zenpan_h)
    zenpan_w = st.number_input("材料規格 全判左右 (mm)", min_value=0.0, value=st.session_state.zenpan_w)

# --- 部材一覧表（） ---
app29_display_rows = []
if st.session_state.loaded_data and st.session_state.app29_parts_rows:
    st.markdown("---")
    st.subheader("📋 ：部材一覧表（）")
    st.caption("💡 【PDF掲載】のチェックでPDF出力の有無を選択でき、項目【内容】のみ画面上で編集可能です。")
    
    df_app29_input = pd.DataFrame(st.session_state.app29_parts_rows)
    edited_app29 = st.data_editor(
        df_app29_input,
        use_container_width=True,
        hide_index=True,
        disabled=["種目"]
    )
    app29_display_rows = edited_app29.to_dict('records')

st.markdown("---")
st.subheader("3. 計算用パーツリスト（アプリ50 独立データ）")
df_parts_input = pd.DataFrame(st.session_state.parts_rows) if st.session_state.parts_rows else pd.DataFrame(columns=["部材名", "銘柄", "備考", "１冊面数", "天地サイズ+-", "左右サイズ+-", "固定幅【背】", "レコードID"])
edited_parts = st.data_editor(df_parts_input, num_rows="dynamic", use_container_width=True)

if st.button("🖨️ 発注書PDF ＆ Excel集計表を生成する", type="primary"):
    if not st.session_state.loaded_data and order_id:
        fetch_and_store_pocket_data(order_id)

    if edited_parts.empty or (len(edited_parts) == 1 and str(edited_parts.iloc[0]["部材名"]) == ""):
        st.warning("⚠️ 計算用部材明細が空です。データを読み込むか明細を追加してください。")
    else:
        try:
            calculated_rows = []
            total_sheets_sum = 0
            total_pieces_sum = 0
            
            for _, row in edited_parts.iterrows():
                p_name = str(row.get("部材名", ""))
                if not p_name or p_name == "nan":
                    continue
                p_brand = str(row.get("銘柄", "")) if pd.notna(row.get("銘柄", "")) else ""
                p_remark = str(row.get("備考", "")) if pd.notna(row.get("備考", "")) else ""
                is_groove = "溝あき" in p_name
                p_count = 0 if is_groove else safe_int(row.get("１冊面数", 1), 1)
                t_plus = safe_float(row.get("天地サイズ+-", 0.0), 0.0)
                s_plus = safe_float(row.get("左右サイズ+-", 0.0), 0.0)
                fixed_w = safe_float(row.get("固定幅【背】", 0.0), 0.0)
                rec_id = safe_int(row.get("レコードID", 0), 0)
                
                w = fixed_w if fixed_w > 0 else (hon_w + s_plus)
                h = hon_h + t_plus
                
                mentori1 = int(zenpan_w // w) * int(zenpan_h // h) if (w > 0 and h > 0) else 0
                mentori2 = int(zenpan_w // h) * int(zenpan_h // w) if (w > 0 and h > 0) else 0
                max_mentori = max(mentori1, mentori2)
                
                pieces_needed = quantity if is_groove else (quantity * p_count)
                sheets_needed = math.ceil(pieces_needed / max_mentori) if max_mentori > 0 else 0
                
                total_sheets_sum += sheets_needed
                total_pieces_sum += pieces_needed
                
                calculated_rows.append({
                    "部材名": p_name, "銘柄": p_brand, "備考": p_remark,
                    "面数": "-" if is_groove else f"{p_count}面",
                    "サイズ_W": w, "サイズ_H": h, "面取り数": max_mentori,
                    "仕上枚数": pieces_needed, "全判枚数": sheets_needed,
                    "レコードID": rec_id, "is_groove": is_groove
                })

            # --- PDF生成 ---
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
            story = []
            
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontName=font_name, fontSize=22, leading=26, alignment=1, textColor=colors.HexColor('#111'))
            
            # 受注番号と書名を16ポイント太字にするスタイル
            header_bold_style = ParagraphStyle('HeaderBoldStyle', parent=styles['Normal'], fontName=font_name, fontSize=16, leading=20, alignment=1, textColor=colors.HexColor('#111'))
            
            due_style = ParagraphStyle('DueStyle', parent=styles['Normal'], fontName=font_name, fontSize=13, leading=17, alignment=1, textColor=colors.HexColor('#c00'))
            normal_style = ParagraphStyle('NormalStyle', parent=styles['Normal'], fontName=font_name, fontSize=9, leading=13)
            center_style = ParagraphStyle('CenterStyle', parent=styles['Normal'], fontName=font_name, fontSize=9, leading=13, alignment=1)
            table_header_style = ParagraphStyle('TableHeaderStyle', parent=styles['Normal'], fontName=font_name, fontSize=9.5, leading=13, alignment=1, textColor=colors.HexColor('#ffffff'))
            
            story.append(Paragraph("<b>発 注 書</b>", title_style))
            story.append(Spacer(1, 6))
            
            # 受注番号と書名を16pt太字で表示
            story.append(Paragraph(f"<b>受注番号：{order_id} ｜ 書名：{book_title}</b>", header_bold_style))
            story.append(Spacer(1, 8))
            
            story.append(Paragraph(f"<b>【 表紙貼外注先 】</b> {supplier_cover} 様 ｜ <font color='#c00'><b>納期： {due_cover} ({ampm_cover})</b></font>", due_style))
            story.append(Spacer(1, 4))
            story.append(Paragraph(f"<b>【 表紙箔押外注先 】</b> {supplier_foil} 様 ｜ <font color='#c00'><b>納期： {due_foil} ({ampm_foil})</b></font>", due_style))
            story.append(Spacer(1, 8))
            story.append(Paragraph(f"<b>納品先:</b> {delivery_to} 様 ｜ <b>必要部数:</b> {quantity}部 ｜ <b>全判規格:</b> {zenpan_w}mm × {zenpan_h}mm ｜ <b>本文サイズ:</b> {hon_w}mm × {hon_h}mm", normal_style))
            story.append(Spacer(1, 10))
            
            # --- 図解セクション ---
            story.append(Paragraph("<b>■ 部材ビジュアル図解 ＆ 計算仕様情報</b>", normal_style))
            story.append(Spacer(1, 4))
            
            card_metrics = []
            for p in calculated_rows:
                p_lower = p["部材名"]
                if p["is_groove"]:
                    card_metrics.append({"w": int(85 * 0.7), "h": 170, "type": "groove", "p": p})
                elif "背ボール" in p_lower:
                    card_metrics.append({"w": int(125 * 0.7), "h": 170, "type": "back", "p": p})
                elif "芯ボール" in p_lower:
                    card_metrics.append({"w": int(125 * 0.8), "h": 170, "type": "core", "p": p})
                else:
                    card_metrics.append({"w": 125, "h": 170, "type": "normal", "p": p})
            
            max_row_w = 530
            rows_layout = []
            current_row = []
            current_w = 0
            for item in card_metrics:
                spacing = 10 if current_row else 0
                if current_w + item["w"] + spacing > max_row_w and current_row:
                    rows_layout.append(current_row)
                    current_row = [item]
                    current_w = item["w"]
                else:
                    current_row.append(item)
                    current_w += item["w"] + spacing
            if current_row:
                rows_layout.append(current_row)
                
            drawing_height = max(190, len(rows_layout) * 185 + 15)
            d = Drawing(540, drawing_height)
            
            for r_idx, row_items in enumerate(rows_layout):
                y_offset = drawing_height - (r_idx + 1) * 185 - 10
                start_x = 15
                for item in row_items:
                    p = item["p"]
                    box_w = item["w"]
                    p_lower = p["部材名"]
                    if item["type"] == "groove":
                        d.add(Rect(start_x, y_offset, box_w, 170, fillColor=colors.HexColor('#fff8ee'), strokeColor=colors.HexColor('#d9822b'), strokeWidth=1, strokeDashArray=[2,2]))
                        d.add(String(start_x + 4, y_offset + 148, f"【{p_lower}】", fontName=font_name, fontSize=7.5, fillColor=colors.HexColor('#b75800')))
                        d.add(String(start_x + 4, y_offset + 130, f"仕上寸法:", fontName=font_name, fontSize=5.5, fillColor=colors.HexColor('#666')))
                        d.add(String(start_x + 4, y_offset + 116, f"{p['サイズ_W']}×{p['サイズ_H']}", fontName=font_name, fontSize=6, fillColor=colors.HexColor('#222')))
                        d.add(Line(start_x + 4, y_offset + 106, start_x + box_w - 4, y_offset + 106, strokeColor=colors.HexColor('#f0d0a0'), strokeWidth=0.5))
                        d.add(String(start_x + 4, y_offset + 88, f"銘柄:{p['銘柄']}", fontName=font_name, fontSize=5.5, fillColor=colors.HexColor('#444')))
                        d.add(String(start_x + 4, y_offset + 66, f"仕上:{p['仕上枚数']}枚", fontName=font_name, fontSize=6.5, fillColor=colors.HexColor('#004')))
                        d.add(String(start_x + 4, y_offset + 44, f"全判:{p['全判枚数']}枚", fontName=font_name, fontSize=6.5, fillColor=colors.HexColor('#006')))
                    else:
                        d.add(Rect(start_x, y_offset, box_w, 170, fillColor=colors.HexColor('#fafafa'), strokeColor=colors.HexColor('#333'), strokeWidth=1.2))
                        d.add(String(start_x + 6, y_offset + 148, f"【{p_lower}】", fontName=font_name, fontSize=8, fillColor=colors.HexColor('#000')))
                        d.add(String(start_x + 6, y_offset + 130, f"寸法:{p['サイズ_W']}×{p['サイズ_H']}mm", fontName=font_name, fontSize=6, fillColor=colors.HexColor('#444')))
                        d.add(String(start_x + 6, y_offset + 116, f"銘柄:{p['銘柄']}", fontName=font_name, fontSize=6, fillColor=colors.HexColor('#222')))
                        mini_x = start_x + (box_w - 42) / 2
                        d.add(Rect(mini_x, y_offset + 72, 42, 32, fillColor=colors.HexColor('#ffffff'), strokeColor=colors.HexColor('#004080'), strokeWidth=1))
                        d.add(String(mini_x + 4, y_offset + 82, f"{p['面取り数']}面取", fontName=font_name, fontSize=7, fillColor=colors.HexColor('#c00')))
                        d.add(String(start_x + 6, y_offset + 54, f"仕上: {p['仕上枚数']}枚", fontName=font_name, fontSize=7.5, fillColor=colors.HexColor('#004')))
                        d.add(String(start_x + 6, y_offset + 36, f"全判: {p['全判枚数']}枚", fontName=font_name, fontSize=8, fillColor=colors.HexColor('#006')))
                    start_x += box_w + 10
                    
            story.append(d)
            story.append(Spacer(1, 15))

            # --- ★ 追加・移動：部材一覧表（図解の下部・チェックが入っているものだけ掲載） ---
            checked_app29_rows = [ar for ar in app29_display_rows if ar.get("PDF掲載", False)]
            if checked_app29_rows:
                story.append(Paragraph("<b>■ 部材一覧表（材料部材明細）</b>", normal_style))
                story.append(Spacer(1, 4))
                app29_table_data = [[Paragraph("<b>種目</b>", table_header_style), Paragraph("<b>内容</b>", table_header_style)]]
                for ar in checked_app29_rows:
                    app29_table_data.append([
                        Paragraph(str(ar.get("種目", "")), normal_style),
                        Paragraph(str(ar.get("銘柄", "")), normal_style)
                    ])
                t_app29 = Table(app29_table_data, colWidths=[200, 340])
                t_app29.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1e3d59')),
                    ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#999')),
                    ('PADDING', (0,0), (-1,-1), 4),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ]))
                story.append(t_app29)
                story.append(Spacer(1, 15))
            
            # --- テーブル構築 ---
            table_data = [[
                Paragraph("<b>部材名</b>", table_header_style), Paragraph("<b>銘柄(表紙クロス)</b>", table_header_style), 
                Paragraph("<b>備考</b>", table_header_style), Paragraph("<b>1冊面数</b>", table_header_style), 
                Paragraph("<b>計算サイズ (W×H)</b>", table_header_style), Paragraph("<b>全判サイズ</b>", table_header_style), 
                Paragraph("<b>面取り数</b>", table_header_style), Paragraph("<b>仕上枚数</b>", table_header_style), 
                Paragraph("<b>全判枚数</b>", table_header_style)
            ]]
            for p in calculated_rows:
                table_data.append([
                    Paragraph(p["部材名"], normal_style), Paragraph(p["銘柄"], normal_style), Paragraph(p["備考"], normal_style),
                    Paragraph(str(p["面数"]), center_style), Paragraph(f"{p['サイズ_W']}mm × {p['サイズ_H']}mm", normal_style),
                    Paragraph(f"{zenpan_w}mm × {zenpan_h}mm", normal_style), Paragraph("-" if p["is_groove"] else f"{p['面取り数']}面/枚", center_style),
                    Paragraph(f"{p['仕上枚数']}枚", center_style), Paragraph(f"<b>{p['全判枚数']}枚</b>", center_style),
                ])
            table_data.append([
                Paragraph("<b>合計</b>", center_style), Paragraph("", normal_style), Paragraph("", normal_style),
                Paragraph("", center_style), Paragraph("", normal_style), Paragraph("", center_style),
                Paragraph("", center_style), Paragraph(f"<b>{total_pieces_sum}枚</b>", center_style), Paragraph(f"<b>{total_sheets_sum}枚</b>", center_style),
            ])
                
            t = Table(table_data, colWidths=[70, 75, 55, 45, 85, 65, 50, 47, 48])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1e3d59')),
                ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#f0f0f0')),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#999')),
                ('PADDING', (0,0), (-1,-1), 4),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ]))
            story.append(t)
            
            # --- 社判・担当者サイン欄の追加 ---
            story.append(Spacer(1, 15))
            sign_header_style = ParagraphStyle('SignHeader', parent=styles['Normal'], fontName=font_name, fontSize=8, alignment=1)
            sign_cell_style = ParagraphStyle('SignCell', parent=styles['Normal'], fontName=font_name, fontSize=8, alignment=1)
            
            sign_data = [
                [
                    Paragraph("<b>発注元</b>", sign_header_style),
                    Paragraph("<b>承認者</b>", sign_header_style),
                    Paragraph("<b>担当者</b>", sign_header_style),
                    Paragraph("<b>受注確認（外注先様）</b>", sign_header_style)
                ],
                [
                    Paragraph("<br/><br/>（社判）<br/><br/>", sign_cell_style),
                    Paragraph("<br/><br/><br/>", sign_cell_style),
                    Paragraph("<br/><br/><br/>", sign_cell_style),
                    Paragraph("<br/><br/><br/>", sign_cell_style)
                ]
            ]
            sign_table = Table(sign_data, colWidths=[100, 80, 80, 280])
            sign_table.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#f0f0f0')),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#999')),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ]))
            story.append(sign_table)
            
            doc.build(story)
            pdf_bytes = buffer.getvalue()
            
            # --- Excel 生成 ---
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "発注計算集計表"
            ws.append(["受注番号", order_id, "書名", book_title])
            ws.append(["必要部数", quantity, "全判サイズ", f"{zenpan_w} × {zenpan_h} mm"])
            ws.append([])
            ws.append(["部材名", "銘柄", "備考", "面数", "サイズW", "サイズH", "面取り数", "仕上枚数", "全判枚数"])
            for p in calculated_rows:
                ws.append([p["部材名"], p["銘柄"], p["備考"], p["面数"], p["サイズ_W"], p["サイズ_H"], p["面取り数"], p["仕上枚数"], p["全判枚数"]])
            ws.append(["合計", "", "", "", "", "", "", total_pieces_sum, total_sheets_sum])
            
            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            excel_bytes = excel_buffer.getvalue()
            
            st.success("✨ 発注書PDF および Excelの生成が完了しました！")
            
            col_dl1, col_dl2 = st.columns(2)
            with col_dl1:
                st.download_button("📥 発注書PDFをダウンロード", data=pdf_bytes, file_name=f"order_sheet_{order_id}.pdf", mime="application/pdf")
            with col_dl2:
                st.download_button("📊 集計Excelをダウンロード", data=excel_bytes, file_name=f"order_calc_{order_id}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                
        except Exception as err:
            st.error(f"ファイル生成中にエラーが発生しました: {err}")