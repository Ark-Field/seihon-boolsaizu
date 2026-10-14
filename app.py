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
st.set_page_config(page_title="芯ボール・背ボール専用 発注計算・図面自動作成システム", layout="wide")

st.title("📦 芯ボール・背ボール専用 発注計算 ＆ 図面付き発注書生成システム")
st.markdown("💡 このシステムは製本用の芯ボール・背ボールなどの部材寸法計算、全判からの面取り数算出、および図面付き発注書PDFの作成を単独で行う独立型ツールです。")

# ── 安全な数値変換 ──
def safe_int(val, default=1):
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

# ── メイン入力フォーム ──
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. 受注・製品基本情報")
    order_id = st.text_input("受注番号", value="ORD-2026-001")
    customer_name = st.text_input("顧客名", value="〇〇出版")
    product_name = st.text_input("製品名（品名）", value="ハードカバー上製本")
    quantity = st.number_input("発行部数", min_value=1, value=500)

with col2:
    st.subheader("2. 本文仕様 ＆ 全判サイズ設定")
    hon_h = st.number_input("本文 天地サイズ (mm)", min_value=10.0, value=297.0)
    hon_w = st.number_input("本文 左右サイズ (mm)", min_value=10.0, value=210.0)
    tenchi_plus = st.number_input("天地 プラスサイズ (mm)", min_value=0.0, value=20.0)
    sayu_plus = st.number_input("左右 プラスサイズ (mm)", min_value=0.0, value=15.0)
    groove_size = st.number_input("溝あきサイズ (背芯幅・mm)", min_value=0.0, value=15.0)
    zenpan_h = st.number_input("全判 天地サイズ (mm)", min_value=100.0, value=700.0)
    zenpan_w = st.number_input("全判 左右サイズ (mm)", min_value=100.0, value=1000.0)

st.markdown("---")
st.subheader("3. マルチ部材明細（パーツリスト設定）")
st.markdown("💡 芯ボール（表表紙・裏表紙）や背ボール、その他必要に応じて行を追加・編集できます。「固定幅(mm)」に値を設定すると、その幅が優先されます（0の場合は本文サイズ＋左右プラスで計算）。")

# 初期マルチ部材データ
default_parts = pd.DataFrame([
    {"部材名": "芯ボール (表表紙)", "1冊面数": 1, "プラス左右(mm)": sayu_plus, "プラス天地(mm)": tenchi_plus, "固定幅(mm)": 0.0},
    {"部材名": "芯ボール (裏表紙)", "1冊面数": 1, "プラス左右(mm)": sayu_plus, "プラス天地(mm)": tenchi_plus, "固定幅(mm)": 0.0},
    {"部材名": "背ボール (背芯)", "1冊面数": 1, "プラス左右(mm)": 0.0, "プラス天地(mm)": tenchi_plus, "固定幅(mm)": groove_size},
])

edited_parts = st.data_editor(default_parts, num_rows="dynamic", use_container_width=True)

if st.button("🖨️ マルチ計算 ＆ 図面付き発注書PDFを生成する", type="primary"):
    try:
        calculated_rows = []
        for _, row in edited_parts.iterrows():
            p_name = str(row["部材名"])
            p_count = safe_int(row["1冊面数"], 1)
            p_plus_w = safe_float(row["プラス左右(mm)"], 0.0)
            p_plus_h = safe_float(row["プラス天地(mm)"], 0.0)
            p_fixed_w = safe_float(row["固定幅(mm)"], 0.0)
            
            # 寸法算出
            w = p_fixed_w if p_fixed_w > 0 else (hon_w + p_plus_w)
            h = hon_h + p_plus_h
            
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
                "面数": p_count,
                "サイズ_W": w,
                "サイズ_H": h,
                "面取り数": max_mentori,
                "必要枚数": sheets_needed
            })

        # PDF生成
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
        story = []
        
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontName=font_name, fontSize=14, leading=18)
        normal_style = ParagraphStyle('NormalStyle', parent=styles['Normal'], fontName=font_name, fontSize=9, leading=13)
        
        story.append(Paragraph(f"<b>【材料発注書・図面付き】 {product_name} (受注: {order_id})</b>", title_style))
        story.append(Spacer(1, 8))
        story.append(Paragraph(f"<b>顧客名:</b> {customer_name} 様 ｜ <b>発注部数:</b> {quantity}冊 ｜ <b>全判サイズ:</b> {zenpan_w}mm × {zenpan_h}mm", normal_style))
        story.append(Spacer(1, 12))
        
        # 図解（パーツ配置図）の描画
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
        table_data = [["部材名", "1冊面数", "計算サイズ (W×H)", "全判サイズ", "面取り数", "発注必要枚数"]]
        for p in calculated_rows:
            table_data.append([
                Paragraph(p["部材名"], normal_style),
                Paragraph(f"{p['面数']}面", normal_style),
                Paragraph(f"{p['サイズ_W']}mm × {p['サイズ_H']}mm", normal_style),
                Paragraph(f"{zenpan_w}mm × {zenpan_h}mm", normal_style),
                Paragraph(f"{p['面取り数']}面 / 枚", normal_style),
                Paragraph(f"<b>{p['必要枚数']}枚</b>", normal_style),
            ])
            
        t = Table(table_data, colWidths=[110, 65, 115, 95, 70, 75])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#e0e0e0')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#999')),
            ('PADDING', (0,0), (-1,-1), 6),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))
        story.append(t)
        
        doc.build(story)
        pdf_bytes = buffer.getvalue()
        
        st.success("✨ マルチ部材の発注計算および図面付きPDFの生成が完了しました！")
        st.download_button(
            label="📥 【図面付き】材料発注書PDFをダウンロードする",
            data=pdf_bytes,
            file_name=f"multi_board_order_{order_id}.pdf",
            mime="application/pdf"
        )
        
    except Exception as pdf_err:
        st.error(f"PDF生成中にエラーが発生しました: {pdf_err}")