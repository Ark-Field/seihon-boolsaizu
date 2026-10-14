import streamlit as st
import pandas as pd
import io
import os
import math
import urllib.request
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.graphics.shapes import Drawing, Rect, String
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

st.set_page_config(page_title="芯ボール・背ボール専用 発注計算・図面自動作成システム", layout="wide")

st.title("📦 芯ボール・背ボール専用 発注計算 ＆ 図面付き発注書生成システム")
st.markdown("💡 ご提示いただいた確定版CSV（@pocket出力フォーマット）をそのまま読み込み、自動で寸法計算・全判面取り数算出・図面付き発注書PDFを生成します。")

# ── 安全な数値変換 ──
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

# ── 日本語フォント登録 ──
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

# ── サイドバー：CSVアップロード ──
with st.sidebar:
    st.header("📂 確定版CSVデータ読込")
    uploaded_file = st.file_uploader("@pocket出力CSVを選択", type=["csv"], help="受注番号CO,書名,必要数量...,部材名,天地サイズ+-,左右サイズ+-,固定幅【背】等を含むCSV")
    
    st.markdown("---")
    st.markdown("### ℹ️ 対応CSV列構成")
    st.code("受注番号CO, 書名, 必要数量（予備含む）, 表紙貼外注先, 納期（表紙）, サイズ判型, 天地, 左右, 材料規格サイズ, 材天地, 材左右, 納品先, 柄名, 部材名, 備考, １冊面数, 天地サイズ+-, 左右サイズ+-, 固定幅【背】, レコードID")

# デフォルト値（サンプルデータ）
order_id = "260543"
book_title = "詩画集①"
quantity = 1060
supplier = "㈱博勝堂"
due_date = "2026/10/14"
delivery_to = "㈱博勝堂"
hon_h = 182.0
hon_w = 128.0
zenpan_h = 1100.0
zenpan_w = 800.0

raw_parts_data = [
    {"部材名": "芯ボール", "備考": "", "１冊面数": 1, "天地サイズ+-": 6.0, "左右サイズ+-": -3.5, "固定幅【背】": 0.0, "レコードID": 1},
    {"部材名": "背ボール", "備考": "", "１冊面数": 1, "天地サイズ+-": 6.0, "左右サイズ+-": 0.0, "固定幅【背】": 12.0, "レコードID": 2},
    {"部材名": "芯ボール", "備考": "", "１冊面数": 1, "天地サイズ+-": 6.0, "左右サイズ+-": -3.5, "固定幅【背】": 0.0, "レコードID": 3},
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
            order_id = str(first.get("受注番号CO", order_id))
            book_title = str(first.get("書名", book_title))
            quantity = safe_int(first.get("必要数量（予備含む）", quantity), quantity)
            supplier = str(first.get("表紙貼外注先", supplier))
            due_date = str(first.get("納期（表紙）", due_date))
            delivery_to = str(first.get("納品先", delivery_to))
            hon_h = safe_float(first.get("天地", hon_h), hon_h)
            hon_w = safe_float(first.get("左右", hon_w), hon_w)
            zenpan_h = safe_float(first.get("材天地", zenpan_h), zenpan_h)
            zenpan_w = safe_float(first.get("材左右", zenpan_w), zenpan_w)
            
            # 明細行の抽出（溝あき等を除外または保持）
            extracted_parts = []
            for _, row in df_csv.iterrows():
                b_name = str(row.get("部材名", "")).strip()
                if not b_name or b_name == "nan" or "溝あき" in b_name:
                    continue
                extracted_parts.append({
                    "部材名": b_name,
                    "備考": str(row.get("備考", "")) if pd.notna(row.get("備考", "")) else "",
                    "１冊面数": safe_int(row.get("１冊面数", 1), 1),
                    "天地サイズ+-": safe_float(row.get("天地サイズ+-", 0.0), 0.0),
                    "左右サイズ+-": safe_float(row.get("左右サイズ+-", 0.0), 0.0),
                    "固定幅【背】": safe_float(row.get("固定幅【背】", 0.0), 0.0),
                    "レコードID": safe_int(row.get("レコードID", 0), 0)
                })
            if extracted_parts:
                raw_parts_data = extracted_parts
                
        st.sidebar.success("✅ 確定版CSVファイルを正常に読み込みました！")
    except Exception as e:
        st.sidebar.error(f"CSV読込エラー: {e}")

# ── メイン入力フォーム ──
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. 受注・発注基本情報")
    order_id = st.text_input("受注番号CO", value=order_id)
    book_title = st.text_input("書名（製品名）", value=book_title)
    quantity = st.number_input("必要数量（予備含む）", min_value=1, value=quantity)
    supplier = st.text_input("表紙貼外注先（発注先）", value=supplier)
    due_date = st.text_input("納期", value=due_date)
    delivery_to = st.text_input("納品先", value=delivery_to)

with col2:
    st.subheader("2. 本文仕様 ＆ 全判材料サイズ設定")
    hon_h = st.number_input("本文 天地サイズ (mm)", min_value=10.0, value=hon_h)
    hon_w = st.number_input("本文 左右サイズ (mm)", min_value=10.0, value=hon_w)
    zenpan_h = st.number_input("材料規格 全判天地 (mm)", min_value=100.0, value=zenpan_h)
    zenpan_w = st.number_input("材料規格 全判左右 (mm)", min_value=100.0, value=zenpan_w)

st.markdown("---")
st.subheader("3. 読込済み部材明細（パーツリスト確認・編集）")
st.markdown("💡 CSVから自動展開された明細データです。必要に応じてテーブル上で直接修正・追加できます。")

df_parts_input = pd.DataFrame(raw_parts_data)
edited_parts = st.data_editor(df_parts_input, num_rows="dynamic", use_container_width=True)

if st.button("🖨️ マルチ計算 ＆ 図面付き発注書PDFを生成する", type="primary"):
    try:
        calculated_rows = []
        for _, row in edited_parts.iterrows():
            p_name = str(row.get("部材名", "芯ボール"))
            p_remark = str(row.get("備考", "")) if pd.notna(row.get("備考", "")) else ""
            p_count = safe_int(row.get("１冊面数", 1), 1)
            t_plus = safe_float(row.get("天地サイズ+-", 0.0), 0.0)
            s_plus = safe_float(row.get("左右サイズ+-", 0.0), 0.0)
            fixed_w = safe_float(row.get("固定幅【背】", 0.0), 0.0)
            rec_id = safe_int(row.get("レコードID", 0), 0)
            
            # 寸法算出
            if fixed_w > 0:
                w = fixed_w
            else:
                w = hon_w + s_plus
                
            h = hon_h + t_plus
            
            # 面取り数計算（縦横の取り合わせ2パターンを比較し大きい方を採用）
            cols1 = int(zenpan_w // w) if w > 0 else 0
            rows1 = int(zenpan_h // h) if h > 0 else 0
            mentori1 = cols1 * rows1
            
            cols2 = int(zenpan_w // h) if h > 0 else 0
            rows2 = int(zenpan_h // w) if h > 0 else 0
            mentori2 = cols2 * rows2
            
            max_mentori = max(mentori1, mentori2)
            total_pieces = quantity * p_count
            sheets_needed = math.ceil(total_pieces / max_mentori) if max_mentori > 0 else 0
            
            calculated_rows.append({
                "部材名": p_name,
                "備考": p_remark,
                "面数": p_count,
                "サイズ_W": w,
                "サイズ_H": h,
                "面取り数": max_mentori,
                "必要枚数": sheets_needed,
                "レコードID": rec_id
            })

        # PDF生成
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
        story = []
        
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontName=font_name, fontSize=14, leading=18)
        normal_style = ParagraphStyle('NormalStyle', parent=styles['Normal'], fontName=font_name, fontSize=9, leading=13)
        
        story.append(Paragraph(f"<b>【材料発注書・図面付き】 {book_title} (受注番号: {order_id})</b>", title_style))
        story.append(Spacer(1, 6))
        story.append(Paragraph(f"<b>発注先:</b> {supplier} 様 ｜ <b>納品先:</b> {delivery_to} 様 ｜ <b>納期:</b> {due_date}", normal_style))
        story.append(Paragraph(f"<b>必要部数:</b> {quantity}冊 ｜ <b>全判規格サイズ:</b> {zenpan_w}mm × {zenpan_h}mm ｜ <b>本文サイズ:</b> {hon_w}mm × {hon_h}mm", normal_style))
        story.append(Spacer(1, 10))
        
        # 図解
        d = Drawing(450, 120)
        start_x = 30
        for i, p in enumerate(calculated_rows):
            box_w = min(110, max(45, int(p['サイズ_W'] * 0.4)))
            d.add(Rect(start_x, 15, box_w, 90, fillColor=colors.HexColor('#f5f5f5' if i%2==0 else '#e0e0e0'), strokeColor=colors.HexColor('#333'), strokeWidth=1))
            d.add(String(start_x + 5, 70, p['部材名'][:10], fontName=font_name, fontSize=8))
            d.add(String(start_x + 5, 50, f"{p['サイズ_W']}×{p['サイズ_H']}mm", fontName=font_name, fontSize=7, fillColor=colors.HexColor('#666')))
            start_x += box_w + 10
            if start_x > 400:
                break
        story.append(d)
        story.append(Spacer(1, 15))
        
        # テーブル構築
        table_data = [["部材名", "備考", "1冊面数", "計算サイズ (W×H)", "全判サイズ", "面取り数", "発注必要枚数"]]
        for p in calculated_rows:
            table_data.append([
                Paragraph(p["部材名"], normal_style),
                Paragraph(p["備考"], normal_style),
                Paragraph(f"{p['面数']}面", normal_style),
                Paragraph(f"{p['Size_W'] if 'Size_W' in p else p['サイズ_W']}mm × {p['サイズ_H']}mm", normal_style), # Safe fallback
                Paragraph(f"{zenpan_w}mm × {zenpan_h}mm", normal_style),
                Paragraph(f"{p['面取り数']}面 / 枚", normal_style),
                Paragraph(f"<b>{p['必要枚数']}枚</b>", normal_style),
            ])
            
        t = Table(table_data, colWidths=[90, 70, 55, 105, 85, 60, 65])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#e0e0e0')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#999')),
            ('PADDING', (0,0), (-1,-1), 6),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))
        story.append(t)
        
        doc.build(story)
        pdf_bytes = buffer.getvalue()
        
        st.success("✨ 確定版CSVに基づくマルチ部材計算および図面付き発注書PDFの生成が完了しました！")
        st.download_button(
            label="📥 【図面付き】材料発注書PDFをダウンロードする",
            data=pdf_bytes,
            file_name=f"order_sheet_{order_id}.pdf",
            mime="application/pdf"
        )
        
    except Exception as pdf_err:
        st.error(f"PDF生成中にエラーが発生しました: {pdf_err}")