import streamlit as st
import pandas as pd
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

# ページ基本設定
st.set_page_config(page_title="芯ボール・背ボール・溝あき対応 発注書・図面自動作成システム", layout="wide")

st.title("📦 芯ボール・背ボール・溝あき対応 発注計算 ＆ 図面付き発注書生成システム")
st.markdown("💡 確定版CSV（@pocket出力フォーマット）の「溝あき」を含む全部材を正確に読み込み、自動で寸法計算・全判面取り数・全判必要枚数・仕上枚数を算出し、ご要望を反映した発注書PDFを生成する独立型アプリケーションです。")

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

with st.sidebar:
    st.header("📂 確定版CSVデータ読込")
    uploaded_file = st.file_uploader("@pocket出力CSVを選択", type=["csv"], help="受注番号CO,書名,必要数量...,納期（表紙）,AMPM（表紙）,納期（箔押）,AMPM（箔押）,柄名(銘柄)等を含むCSV")

default_order_id = "260543"
default_book_title = "詩画集①"
default_quantity = 1060
default_supplier_cover = "㈱博勝堂"
default_due_cover = "2026-10-14"
default_ampm_cover = "AM"
default_supplier_foil = "丸栄有限会社"
default_due_foil = "2026-10-14"
default_ampm_foil = "AM"
default_delivery_to = "㈱博勝堂"
default_hon_h = 182.0
default_hon_w = 128.0
default_zenpan_h = 1100.0
default_zenpan_w = 800.0

default_parts_data = [
    {"部材名": "芯ボール", "銘柄": "NPCC#26", "備考": "", "１冊面数": 1, "天地サイズ+-": 6.0, "左右サイズ+-": -3.5, "固定幅【背】": 0.0, "レコードID": 1},
    {"部材名": "溝あき", "銘柄": "", "備考": "", "１冊面数": 0, "天地サイズ+-": 0.0, "左右サイズ+-": 0.0, "固定幅【背】": 0.0, "レコードID": 4},
    {"部材名": "背ボール", "銘柄": "NPCC#26", "備考": "", "１冊面数": 1, "天地サイズ+-": 6.0, "左右サイズ+-": 0.0, "固定幅【背】": 12.0, "レコードID": 2},
    {"部材名": "溝あき", "銘柄": "", "備考": "", "１冊面数": 0, "天地サイズ+-": 0.0, "左右サイズ+-": 0.0, "固定幅【背】": 0.0, "レコードID": 5},
    {"部材名": "芯ボール", "銘柄": "NPCC#26", "備考": "", "１冊面数": 1, "天地サイズ+-": 6.0, "左右サイズ+-": -3.5, "固定幅【背】": 0.0, "レコードID": 3},
]

if uploaded_file is not None:
    try:
        csv_bytes = uploaded_file.read()
        text = ""
        for enc in ['utf-8', 'cp932', 'shift_jis']:
            try:
                text = csv_bytes.decode(enc)
                break
            except:
                continue
        
        df_csv = pd.read_csv(io.StringIO(text))
        df_csv.columns = df_csv.columns.str.strip()
        
        if not df_csv.empty:
            first = df_csv.iloc[0]
            default_order_id = str(first.get("受注番号CO", default_order_id))
            default_book_title = str(first.get("書名", default_book_title))
            default_quantity = safe_int(first.get("必要数量（予備含む）", default_quantity), default_quantity)
            default_supplier_cover = str(first.get("表紙貼外注先", default_supplier_cover))
            default_due_cover = str(first.get("納期（表紙）", default_due_cover))
            default_ampm_cover = str(first.get("AMPM（表紙）", default_ampm_cover))
            default_supplier_foil = str(first.get("表紙箔押外注先", default_supplier_foil))
            default_due_foil = str(first.get("納期（箔押）", default_due_foil))
            default_ampm_foil = str(first.get("AMPM（箔押）", default_ampm_foil))
            default_delivery_to = str(first.get("納品先", default_delivery_to))
            default_hon_h = safe_float(first.get("天地", default_hon_h), default_hon_h)
            default_hon_w = safe_float(first.get("左右", default_hon_w), default_hon_w)
            default_zenpan_h = safe_float(first.get("材天地", default_zenpan_h), default_zenpan_h)
            default_zenpan_w = safe_float(first.get("材左右", default_zenpan_w), default_zenpan_w)
            
            extracted_parts = []
            for _, row in df_csv.iterrows():
                b_name = str(row.get("部材名", "")).strip()
                if not b_name or b_name == "nan":
                    continue
                is_groove = "溝あき" in b_name
                extracted_parts.append({
                    "部材名": b_name,
                    "銘柄": str(row.get("柄名", "")) if pd.notna(row.get("柄名", "")) else "",
                    "備考": str(row.get("備考", "")) if pd.notna(row.get("備考", "")) else "",
                    "１冊面数": 0 if is_groove else safe_int(row.get("１冊面数", 1), 1),
                    "天地サイズ+-": safe_float(row.get("天地サイズ+-", 0.0), 0.0),
                    "左右サイズ+-": safe_float(row.get("左右サイズ+-", 0.0), 0.0),
                    "固定幅【背】": safe_float(row.get("固定幅【背】", 0.0), 0.0),
                    "レコードID": safe_int(row.get("レコードID", 0), 0)
                })
            if extracted_parts:
                default_parts_data = extracted_parts
                
        st.sidebar.success("✅ 確定版CSVファイルを正常に読み込みました！（AM/PM納品時間＆柄名対応）")
    except Exception as e:
        st.sidebar.error(f"CSV読込エラー: {e}")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. 受注・発注基本情報")
    order_id = st.text_input("受注番号CO", value=default_order_id)
    book_title = st.text_input("書名（製品名）", value=default_book_title)
    quantity = st.number_input("必要数量（予備含む）", min_value=1, value=default_quantity)
    supplier_cover = st.text_input("表紙貼外注先", value=default_supplier_cover)
    due_cover = st.text_input("納期（表紙）", value=default_due_cover)
    ampm_cover = st.text_input("AMPM（表紙）", value=default_ampm_cover)
    supplier_foil = st.text_input("表紙箔押外注先", value=default_supplier_foil)
    due_foil = st.text_input("納期（箔押）", value=default_due_foil)
    ampm_foil = st.text_input("AMPM（箔押）", value=default_ampm_foil)
    delivery_to = st.text_input("納品先", value=default_delivery_to)

with col2:
    st.subheader("2. 本文仕様 ＆ 全判材料サイズ設定")
    hon_h = st.number_input("本文 天地サイズ (mm)", min_value=10.0, value=default_hon_h)
    hon_w = st.number_input("本文 左右サイズ (mm)", min_value=10.0, value=default_hon_w)
    zenpan_h = st.number_input("材料規格 全判天地 (mm)", min_value=100.0, value=default_zenpan_h)
    zenpan_w = st.number_input("材料規格 全判左右 (mm)", min_value=100.0, value=default_zenpan_w)

st.markdown("---")
st.subheader("3. 読込済み部材明細（パーツリスト確認・編集）")
st.markdown("💡 溝あきは1冊面数を「0」（または空欄）にすることで面数計算から除外できます。")

df_parts_input = pd.DataFrame(default_parts_data)
edited_parts = st.data_editor(df_parts_input, num_rows="dynamic", use_container_width=True)

if st.button("🖨️ 発注書PDFを生成する", type="primary"):
    try:
        calculated_rows = []
        total_sheets_sum = 0
        total_pieces_sum = 0
        
        for _, row in edited_parts.iterrows():
            p_name = str(row.get("部材名", "芯ボール"))
            p_brand = str(row.get("銘柄", "")) if pd.notna(row.get("銘柄", "")) else ""
            p_remark = str(row.get("備考", "")) if pd.notna(row.get("備考", "")) else ""
            is_groove = "溝あき" in p_name
            p_count = 0 if is_groove else safe_int(row.get("１冊面数", 1), 1)
            t_plus = safe_float(row.get("天地サイズ+-", 0.0), 0.0)
            s_plus = safe_float(row.get("左右サイズ+-", 0.0), 0.0)
            fixed_w = safe_float(row.get("固定幅【背】", 0.0), 0.0)
            rec_id = safe_int(row.get("レコードID", 0), 0)
            
            if fixed_w > 0:
                w = fixed_w
            else:
                w = hon_w + s_plus
                
            h = hon_h + t_plus
            
            cols1 = int(zenpan_w // w) if w > 0 else 0
            rows1 = int(zenpan_h // h) if h > 0 else 0
            mentori1 = cols1 * rows1
            
            cols2 = int(zenpan_w // h) if h > 0 else 0
            rows2 = int(zenpan_h // w) if h > 0 else 0
            mentori2 = cols2 * rows2
            
            max_mentori = max(mentori1, mentori2)
            
            if is_groove:
                pieces_needed = quantity
            else:
                pieces_needed = quantity * p_count
                
            sheets_needed = math.ceil(pieces_needed / max_mentori) if max_mentori > 0 else 0
            
            total_sheets_sum += sheets_needed
            total_pieces_sum += pieces_needed
            
            calculated_rows.append({
                "部材名": p_name,
                "銘柄": p_brand,
                "備考": p_remark,
                "面数": "-" if is_groove else f"{p_count}面",
                "サイズ_W": w,
                "サイズ_H": h,
                "面取り数": max_mentori,
                "仕上枚数": pieces_needed,
                "全判枚数": sheets_needed,
                "レコードID": rec_id,
                "is_groove": is_groove
            })

        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
        story = []
        
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontName=font_name, fontSize=22, leading=26, alignment=1, textColor=colors.HexColor('#111'))
        due_style = ParagraphStyle('DueStyle', parent=styles['Normal'], fontName=font_name, fontSize=13, leading=17, alignment=1, textColor=colors.HexColor('#c00'))
        normal_style = ParagraphStyle('NormalStyle', parent=styles['Normal'], fontName=font_name, fontSize=9, leading=13)
        center_style = ParagraphStyle('CenterStyle', parent=styles['Normal'], fontName=font_name, fontSize=9, leading=13, alignment=1)
        
        table_header_style = ParagraphStyle(
            'TableHeaderStyle', 
            parent=styles['Normal'], 
            fontName=font_name, 
            fontSize=9.5, 
            leading=13, 
            alignment=1, 
            textColor=colors.HexColor('#ffffff')
        )
        table_header_center = ParagraphStyle(
            'TableHeaderCenter', 
            parent=table_header_style, 
            alignment=1
        )
        
        # 1行目：大きめで発注書
        story.append(Paragraph("<b>発 注 書</b>", title_style))
        story.append(Spacer(1, 6))
        
        # 納期の後ろに納品時間（AMPM）を表示
        story.append(Paragraph(f"<b>【 表紙貼外注先 】</b> {supplier_cover} 様 ｜ <font color='#c00'><b>納期： {due_cover} ({ampm_cover})</b></font>", due_style))
        story.append(Spacer(1, 4))
        story.append(Paragraph(f"<b>【 表紙箔押外注先 】</b> {supplier_foil} 様 ｜ <font color='#c00'><b>納期： {due_foil} ({ampm_foil})</b></font>", due_style))
        story.append(Spacer(1, 8))
        
        story.append(Paragraph(f"<b>納品先:</b> {delivery_to} 様 ｜ <b>受注番号:</b> {order_id} ｜ <b>書名:</b> {book_title}", normal_style))
        story.append(Paragraph(f"<b>必要部数:</b> {quantity}部 ｜ <b>全判規格:</b> {zenpan_w}mm × {zenpan_h}mm ｜ <b>本文サイズ:</b> {hon_w}mm × {hon_h}mm", normal_style))
        story.append(Spacer(1, 10))
        
        # --- 図解セクション ---
        story.append(Paragraph("<b>■ 部材ビジュアル図解 ＆ 仕様情報（芯ボール20%減・溝あき30%減）</b>", normal_style))
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
            
        num_drawing_rows = len(rows_layout)
        drawing_height = max(190, num_drawing_rows * 185 + 15)
        
        d = Drawing(540, drawing_height)
        
        for r_idx, row_items in enumerate(rows_layout):
            y_offset = drawing_height - (r_idx + 1) * 185 - 10
            start_x = 15
            
            for item in row_items:
                p = item["p"]
                box_w = item["w"]
                box_h = 170
                p_lower = p["部材名"]
                
                if item["type"] == "groove":
                    d.add(Rect(start_x, y_offset, box_w, box_h, fillColor=colors.HexColor('#fff8ee'), strokeColor=colors.HexColor('#d9822b'), strokeWidth=1, strokeDashArray=[2,2]))
                    d.add(String(start_x + 4, y_offset + 148, f"【{p_lower}】", fontName=font_name, fontSize=7.5, fillColor=colors.HexColor('#b75800')))
                    d.add(String(start_x + 4, y_offset + 130, f"仕上寸法:", fontName=font_name, fontSize=5.5, fillColor=colors.HexColor('#666')))
                    d.add(String(start_x + 4, y_offset + 116, f"{p['サイズ_W']}×{p['サイズ_H']}", fontName=font_name, fontSize=6, fillColor=colors.HexColor('#222')))
                    
                    d.add(Line(start_x + 4, y_offset + 106, start_x + box_w - 4, y_offset + 106, strokeColor=colors.HexColor('#f0d0a0'), strokeWidth=0.5))
                    
                    d.add(String(start_x + 4, y_offset + 88, "※数値表示", fontName=font_name, fontSize=5.5, fillColor=colors.HexColor('#888')))
                    d.add(String(start_x + 4, y_offset + 66, f"仕上:{p['仕上枚数']}枚", fontName=font_name, fontSize=6.5, fillColor=colors.HexColor('#004')))
                    d.add(String(start_x + 4, y_offset + 44, f"全判:{p['全判枚数']}枚", fontName=font_name, fontSize=6.5, fillColor=colors.HexColor('#006')))
                    
                elif item["type"] == "back":
                    d.add(Rect(start_x, y_offset, box_w, box_h, fillColor=colors.HexColor('#eaf4fc'), strokeColor=colors.HexColor('#004080'), strokeWidth=1.2))
                    d.add(String(start_x + 5, y_offset + 148, f"【{p_lower}】", fontName=font_name, fontSize=8, fillColor=colors.HexColor('#000')))
                    d.add(String(start_x + 5, y_offset + 130, f"寸法:{p['サイズ_W']}×{p['サイズ_H']}mm", fontName=font_name, fontSize=6, fillColor=colors.HexColor('#444')))
                    
                    mini_rect_w = 26
                    mini_rect_h = 35
                    mini_x = start_x + (box_w - mini_rect_w) / 2
                    mini_y = y_offset + 82
                    d.add(Rect(mini_x, mini_y, mini_rect_w, mini_rect_h, fillColor=colors.HexColor('#ffffff'), strokeColor=colors.HexColor('#004080'), strokeWidth=1))
                    d.add(String(mini_x + 2, mini_y + 13, f"{p['面取り数']}面", fontName=font_name, fontSize=6, fillColor=colors.HexColor('#c00')))
                    
                    d.add(String(start_x + 5, y_offset + 64, f"仕上: {p['仕上枚数']}枚", fontName=font_name, fontSize=7, fillColor=colors.HexColor('#004')))
                    d.add(String(start_x + 5, y_offset + 44, f"全判: {p['全判枚数']}枚", fontName=font_name, fontSize=7.5, fillColor=colors.HexColor('#006')))
                    
                elif item["type"] == "core":
                    d.add(Rect(start_x, y_offset, box_w, box_h, fillColor=colors.HexColor('#f5f5f5'), strokeColor=colors.HexColor('#333'), strokeWidth=1.2))
                    d.add(String(start_x + 6, y_offset + 148, f"【{p_lower}】", fontName=font_name, fontSize=8.5, fillColor=colors.HexColor('#000')))
                    d.add(String(start_x + 6, y_offset + 130, f"寸法:{p['サイズ_W']}×{p['サイズ_H']}mm", fontName=font_name, fontSize=6.5, fillColor=colors.HexColor('#444')))
                    
                    mini_rect_w = 42
                    mini_rect_h = 35
                    mini_x = start_x + (box_w - mini_rect_w) / 2
                    mini_y = y_offset + 82
                    d.add(Rect(mini_x, mini_y, mini_rect_w, mini_rect_h, fillColor=colors.HexColor('#ffffff'), strokeColor=colors.HexColor('#004080'), strokeWidth=1))
                    d.add(String(mini_x + 4, mini_y + 13, f"{p['面取り数']}面取", fontName=font_name, fontSize=7, fillColor=colors.HexColor('#c00')))
                    
                    d.add(String(start_x + 6, y_offset + 64, f"仕上: {p['仕上枚数']}枚", fontName=font_name, fontSize=7.5, fillColor=colors.HexColor('#004')))
                    d.add(String(start_x + 6, y_offset + 44, f"全判: {p['全判枚数']}枚", fontName=font_name, fontSize=8, fillColor=colors.HexColor('#006')))
                    
                else:
                    d.add(Rect(start_x, y_offset, box_w, box_h, fillColor=colors.HexColor('#fafafa'), strokeColor=colors.HexColor('#333'), strokeWidth=1.2))
                    d.add(String(start_x + 8, y_offset + 148, f"【{p_lower}】", fontName=font_name, fontSize=9, fillColor=colors.HexColor('#000')))
                    d.add(String(start_x + 8, y_offset + 130, f"寸法: {p['サイズ_W']}×{p['サイズ_H']}mm", fontName=font_name, fontSize=7.5, fillColor=colors.HexColor('#444')))
                    
                    mini_rect_w = 50
                    mini_rect_h = 35
                    mini_x = start_x + (box_w - mini_rect_w) / 2
                    mini_y = y_offset + 82
                    d.add(Rect(mini_x, mini_y, mini_rect_w, mini_rect_h, fillColor=colors.HexColor('#ffffff'), strokeColor=colors.HexColor('#004080'), strokeWidth=1))
                    d.add(String(mini_x + 6, mini_y + 13, f"{p['面取り数']}面取", fontName=font_name, fontSize=7.5, fillColor=colors.HexColor('#c00')))
                    
                    d.add(String(start_x + 8, y_offset + 64, f"仕上枚数: {p['仕上枚数']}枚", fontName=font_name, fontSize=8, fillColor=colors.HexColor('#004')))
                    d.add(String(start_x + 8, y_offset + 44, f"全判枚数: {p['全判枚数']}枚", fontName=font_name, fontSize=9, fillColor=colors.HexColor('#006')))
                
                start_x += box_w + 10
                
        story.append(d)
        story.append(Spacer(1, 15))
        
        # --- テーブル構築 ---
        table_data = [[
            Paragraph("<b>部材名</b>", table_header_style), 
            Paragraph("<b>銘柄(表紙クロス)</b>", table_header_style), 
            Paragraph("<b>備考</b>", table_header_style), 
            Paragraph("<b>1冊面数</b>", table_header_center), 
            Paragraph("<b>計算サイズ (W×H)</b>", table_header_style), 
            Paragraph("<b>全判サイズ</b>", table_header_style), 
            Paragraph("<b>面取り数</b>", table_header_center), 
            Paragraph("<b>仕上枚数</b>", table_header_center), 
            Paragraph("<b>全判枚数</b>", table_header_center)
        ]]
        
        for p in calculated_rows:
            table_data.append([
                Paragraph(p["部材名"], normal_style),
                Paragraph(p["銘柄"], normal_style),
                Paragraph(p["備考"], normal_style),
                Paragraph(str(p["面数"]), center_style),
                Paragraph(f"{p['サイズ_W']}mm × {p['サイズ_H']}mm", normal_style),
                Paragraph(f"{zenpan_w}mm × {zenpan_h}mm", normal_style),
                Paragraph("-" if p["is_groove"] else f"{p['面取り数']}面/枚", center_style),
                Paragraph(f"{p['仕上枚数']}枚", center_style),
                Paragraph(f"<b>{p['全判枚数']}枚</b>", center_style),
            ])
            
        # 合計行
        table_data.append([
            Paragraph("<b>合計</b>", center_style),
            Paragraph("", normal_style),
            Paragraph("", normal_style),
            Paragraph("", center_style),
            Paragraph("", normal_style),
            Paragraph("", center_style),
            Paragraph("", center_style),
            Paragraph(f"<b>{total_pieces_sum}枚</b>", center_style),
            Paragraph(f"<b>{total_sheets_sum}枚</b>", center_style),
        ])
            
        t = Table(table_data, colWidths=[70, 75, 55, 45, 85, 65, 50, 47, 48])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2c3e50')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#f0f0f0')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#999')),
            ('PADDING', (0,0), (-1,-1), 4),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))
        story.append(t)
        
        doc.build(story)
        pdf_bytes = buffer.getvalue()
        
        st.success("✨ 納期直後の納品時間（AM/PM）表示対応版発注書PDFの生成が完了しました！")
        st.download_button(
            label="📥 【最新完全版】発注書PDFをダウンロードする",
            data=pdf_bytes,
            file_name=f"order_sheet_{order_id}.pdf",
            mime="application/pdf"
        )
        
    except Exception as pdf_err:
        st.error(f"PDF生成中にエラーが発生しました: {pdf_err}")