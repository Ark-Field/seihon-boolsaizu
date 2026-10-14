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

# ページ基本設定
st.set_page_config(page_title="芯ボール・背ボール・溝あき 発注計算・図面自動作成システム", layout="wide")

st.title("📦 芯ボール・背ボール・溝あき対応 発注計算 ＆ 図面付き発注書生成システム")
st.markdown("💡 確定版CSV（@pocket出力フォーマット）の「溝あき」を含む全部材を正確に読み込み、自動で寸法計算・全判面取り数算出・**大きく見やすい図面付き**発注書PDFを生成する独立型アプリケーションです。")

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

# ── 日本語フォント（NotoSansJP）の自動登録 ──
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
    uploaded_file = st.file_uploader("@pocket出力CSVを選択", type=["csv"], help="受注番号CO,書名,必要数量...,部材名(芯ボール,背ボール,溝あき等),天地サイズ+-,左右サイズ+-,固定幅【背】等を含むCSV")
    
    st.markdown("---")
    st.markdown("### ℹ️ 対応CSV列構成")
    st.code("受注番号CO,書名,必要数量（予備含む）,表紙貼外注先,納期（表紙）,表紙箔押外注先,納期（箔押）,サイズ判型,天地,左右,材料規格サイズ,材天地,材左右,納品先,柄名,部材名,備考,１冊面数,天地サイズ+-,左右サイズ+-,固定幅【背】,レコードID")

# ── デフォルト値（サンプルデータ） ──
default_order_id = "260543"
default_book_title = "詩画集①"
default_quantity = 1060
default_supplier = "㈱博勝堂"
default_due_date = "2026/10/14"
default_delivery_to = "㈱博勝堂"
default_hon_h = 182.0
default_hon_w = 128.0
default_zenpan_h = 1100.0
default_zenpan_w = 800.0

default_parts_data = [
    {"部材名": "芯ボール", "備考": "", "１冊面数": 1, "天地サイズ+-": 6.0, "左右サイズ+-": -3.5, "固定幅【背】": 0.0, "レコードID": 1},
    {"部材名": "溝あき", "備考": "", "１冊面数": 2, "天地サイズ+-": 0.0, "左右サイズ+-": 0.0, "固定幅【背】": 0.0, "レコードID": 2},
    {"部材名": "背ボール", "備考": "", "１冊面数": 1, "天地サイズ+-": 6.0, "左右サイズ+-": 0.0, "固定幅【背】": 12.0, "レコードID": 3},
    {"部材名": "芯ボール", "備考": "", "１冊面数": 1, "天地サイズ+-": 6.0, "左右サイズ+-": -3.5, "固定幅【背】": 0.0, "レコードID": 4},
]

# ── CSVファイル読込・解析処理 ──
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
            default_supplier = str(first.get("表紙貼外注先", default_supplier))
            default_due_date = str(first.get("納期（表紙）", default_due_date))
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
                default_parts_data = extracted_parts
                
        st.sidebar.success("✅ 確定版CSVファイル（溝あき含む）を正常に読み込みました！")
    except Exception as e:
        st.sidebar.error(f"CSV読込エラー: {e}")

# ── メイン入力フォーム ──
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. 受注・発注基本情報")
    order_id = st.text_input("受注番号CO", value=default_order_id)
    book_title = st.text_input("書名（製品名）", value=default_book_title)
    quantity = st.number_input("必要数量（予備含む）", min_value=1, value=default_quantity)
    supplier = st.text_input("表紙貼外注先（発注先）", value=default_supplier)
    due_date = st.text_input("納期", value=default_due_date)
    delivery_to = st.text_input("納品先", value=default_delivery_to)

with col2:
    st.subheader("2. 本文仕様 ＆ 全判材料サイズ設定")
    hon_h = st.number_input("本文 天地サイズ (mm)", min_value=10.0, value=default_hon_h)
    hon_w = st.number_input("本文 左右サイズ (mm)", min_value=10.0, value=default_hon_w)
    zenpan_h = st.number_input("材料規格 全判天地 (mm)", min_value=100.0, value=default_zenpan_h)
    zenpan_w = st.number_input("材料規格 全判左右 (mm)", min_value=100.0, value=default_zenpan_w)

st.markdown("---")
st.subheader("3. 読込済み部材明細（溝あき含むパーツリスト確認・編集）")
st.markdown("💡 CSVから「溝あき」を含めて自動展開された明細データです。必要に応じてテーブル上で直接修正・追加できます。")

df_parts_input = pd.DataFrame(default_parts_data)
edited_parts = st.data_editor(df_parts_input, num_rows="dynamic", use_container_width=True)

if st.button("🖨️ マルチ計算 ＆ 大きな図面付き発注書PDFを生成する", type="primary"):
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
        
        # 図解（図形を大きく表示・各部材ごとに大きく描画）
        story.append(Paragraph("<b>■ 部材寸法・全判面取りレイアウト図解</b>", normal_style))
        story.append(Spacer(1, 4))
        
        d = Drawing(480, 160)
        start_x = 20
        for i, p in enumerate(calculated_rows):
            box_w = min(140, max(80, int(p['サイズ_W'] * 0.5)))
            box_h = 130
            d.add(Rect(start_x, 15, box_w, box_h, fillColor=colors.HexColor('#f9f9f9' if i%2==0 else '#eef2f7'), strokeColor=colors.HexColor('#222'), strokeWidth=1.5))
            
            # テキスト情報を大きく見やすく配置
            d.add(String(start_x + 8, 120, f"【{p['部材名']}】", fontName=font_name, fontSize=10, fillColor=colors.HexColor('#000')))
            d.add(String(start_x + 8, 100, f"仕上寸: {p['サイズ_W']}×{p['サイズ_H']}mm", fontName=font_name, fontSize=9, fillColor=colors.HexColor('#333')))
            d.add(String(start_x + 8, 80, f"1冊面数: {p['面数']}面", fontName=font_name, fontSize=9, fillColor=colors.HexColor('#333')))
            d.add(String(start_x + 8, 55, f"★ 面取り数: {p['面取り数']}面/枚", fontName=font_name, fontSize=10, fillColor=colors.HexColor('#c00')))
            d.add(String(start_x + 8, 30, f"必要枚数: {p['必要枚数']}枚", fontName=font_name, fontSize=10, fillColor=colors.HexColor('#006')))
            
            start_x += box_w + 15
            if start_x > 450:
                break
        story.append(d)
        story.append(Spacer(1, 15))
        
        # テーブル構築（「面取り数」カラムを明記）
        table_data = [["部材名", "備考", "1冊面数", "計算サイズ (W×H)", "全判サイズ", "面取り数", "発注必要枚数"]]
        for p in calculated_rows:
            table_data.append([
                Paragraph(p["部材名"], normal_style),
                Paragraph(p["備考"], normal_style),
                Paragraph(f"{p['面数']}面", normal_style),
                Paragraph(f"{p['サイズ_W']}mm × {p['サイズ_H']}mm", normal_style),
                Paragraph(f"{zenpan_w}mm × {zenpan_h}mm", normal_style),
                Paragraph(f"<b>{p['面取り数']}面 / 枚</b>", normal_style),
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
        
        st.success("✨ 確定版CSV（溝あき含む）に基づくマルチ部材計算および【大きな図面付き】発注書PDFの生成が完了しました！")
        st.download_button(
            label="📥 【大きな図面付き】材料発注書PDFをダウンロードする",
            data=pdf_bytes,
            file_name=f"order_sheet_{order_id}.pdf",
            mime="application/pdf"
        )
        
    except Exception as pdf_err:
        st.error(f"PDF生成中にエラーが発生しました: {pdf_err}")
