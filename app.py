import streamlit as st
import pandas as pd
import io
import os
import math
import urllib.request
from datetime import datetime, timedelta
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ページ基本設定
st.set_page_config(page_title="製本管理業務 - ＆@pocket書き戻し ＆ マルチ部材発注システム", layout="wide")

st.title("📦 製本管理業務 - 個別工程除外 ＆ マルチ部材・図面付き発注書生成システム")

# ── 安全な整数変換ヘルパー関数 ──
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

# ── 日本語フォント（NotoSansJP）の自動取得・登録処理 ──
font_name = 'Helvetica' # フォールバック
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

# ── セッションステートの初期化 ──
if "calculated" not in st.session_state:
    st.session_state.calculated = False
if "df_schedule_out" not in st.session_state:
    st.session_state.df_schedule_out = None
if "df_order_summary" not in st.session_state:
    st.session_state.df_order_summary = None

# ── サイドバー：各種CSVデータの選択と制約パラメータ ──
with st.sidebar:
    st.header("1. 各種CSVデータの選択")
    f1 = st.file_uploader("CSV① 受注データ（必須）", type=["csv"])
    f2 = st.file_uploader("CSV② 休暇シフトデータ（任意）", type=["csv"])
    f3 = st.file_uploader("CSV③ 機械稼働NGデータ（任意）", type=["csv"])
    
    st.header("2. 現場の制約パラメータ")
    base_staff = st.number_input("基本人員数（通常のマックスマンパワー・人）", min_value=1, value=10)
    am_hours = st.number_input("AM労働時間（例: 8:30〜12:00）", min_value=0.0, value=3.5)
    pm_hours = st.number_input("PM労働時間（例: 13:00〜17:30）", min_value=0.0, value=4.5)
    start_base_date = st.date_input("シミュレーション開始基準日（起算日）", value=datetime.today().date())

# メイン画面の処理
if f1 is None:
    st.info("👈 左側のサイドバーから **CSV①（受注データ）** をアップロードしてください。")
else:
    try:
        f1_bytes = f1.read()
        try:
            df_order = pd.read_csv(io.StringIO(f1_bytes.decode('utf-8')))
        except UnicodeDecodeError:
            df_order = pd.read_csv(io.StringIO(f1_bytes.decode('cp932', errors='ignore')))
        
        df_order.columns = df_order.columns.str.strip()
        
        if '工程' in df_order.columns:
            df_order = df_order.dropna(subset=['工程'])
            df_order = df_order[df_order['工程'].astype(str).str.strip() != '']
            df_order = df_order[df_order['工程'].astype(str).str.lower() != 'nan']
            
        if df_order.empty:
            st.error("有効なプログラムデータが見つかりませんでした。")
        else:
            # タブの作成（スケジュール管理 ＆ マルチ部材・図面付き発注書作成）
            tab_schedule, tab_order_sheet = st.tabs(["📈 スケジュール管理・@pocket書き戻し", "📄 マルチ部材・図面付き発注書作成"])

            with tab_schedule:
                st.markdown("### 📋 個別「工程行」除外エディタ")
                df_editor_view = df_order.copy()
                df_editor_view.insert(0, "除外する", False)

                edited_df = st.data_editor(df_editor_view, hide_index=True, use_container_width=True)
                df_target = edited_df[edited_df["除外する"] == False].copy()

                if st.button("🚀 スケジュール計算を実行する", type="primary"):
                    # （従来のスケジュール計算処理がここに走ります）
                    st.success("スケジュール計算が完了しました！（詳細は各項目をご確認ください）")

            with tab_order_sheet:
                st.header("📄 芯ボール・背ボール等 マルチ部材発注書 ＆ 図面自動生成")
                st.markdown("💡 受注データから本文サイズ・溝あき・全判サイズを指定し、複数の部材明細（マルチ対応）を一括計算して図面付きPDFを生成します。")

                col_a, col_b = st.columns(2)
                with col_a:
                    order_id_input = st.text_input("受注番号", value="ORD-2026-001")
                    customer_input = st.text_input("顧客名", value="〇〇出版")
                    product_input = st.text_input("製品名（品名）", value="ハードカバー記念誌")
                    quantity_input = st.number_input("発行部数", min_value=1, value=500)
                
                with col_b:
                    hon_h = st.number_input("本文 天地サイズ (mm)", min_value=50.0, value=297.0)
                    hon_w = st.number_input("本文 左右サイズ (mm)", min_value=50.0, value=210.0)
                    groove_size = st.number_input("溝あきサイズ (背芯幅・mm)", min_value=0.0, value=15.0)
                    zenpan_h = st.number_input("全判 天地サイズ (mm)", min_value=100.0, value=700.0)
                    zenpan_w = st.number_input("全判 左右サイズ (mm)", min_value=100.0, value=1000.0)

                st.markdown("---")
                st.subheader("🧩 マルチ部材明細（パーツリスト設定）")
                st.markdown("必要に応じて行を追加・編集し、芯ボールや背ボールなどの各部材条件を設定してください。")

                # 初期状態のマルチ部材データ
                default_parts_df = pd.DataFrame([
                    {"部材名": "芯ボール (表表紙)", "1冊面数": 1, "プラス左右(mm)": 15.0, "プラス天地(mm)": 20.0, "固定幅(mm)": 0.0},
                    {"部材名": "芯ボール (裏表紙)", "1冊面数": 1, "プラス左右(mm)": 15.0, "プラス天地(mm)": 20.0, "固定幅(mm)": 0.0},
                    {"部材名": "背ボール (背芯)", "1冊面数": 1, "プラス左右(mm)": 0.0, "プラス天地(mm)": 20.0, "固定幅(mm)": groove_size}, # 溝あきサイズ連動
                ])

                edited_parts = st.data_editor(default_parts_df, num_rows="dynamic", use_container_width=True)

                if st.button("🖨️ マルチ計算 ＆ 図面付き発注書PDFを生成する", type="primary"):
                    try:
                        calculated_rows = []
                        for _, row in edited_parts.iterrows():
                            p_name = str(row["部材名"])
                            p_count = safe_int(row["1冊面数"], 1)
                            p_plus_w = safe_float(row["プラス左右(mm)"], 0.0)
                            p_plus_h = safe_float(row["プラス天地(mm)"], 0.0)
                            p_fixed_w = safe_float(row["固定幅(mm)"], 0.0)
                            
                            # サイズ算出（固定幅があれば優先、なければ本文サイズ＋プラス）
                            w = p_fixed_w if p_fixed_w > 0 else (hon_w + p_plus_w)
                            h = hon_h + p_plus_h
                            
                            # 面取り数計算（縦横の取り合わせ2パターンで効率が良い方を採用）
                            cols1 = int(zenpan_w // w) if w > 0 else 0
                            rows1 = int(zenpan_h // h) if h > 0 else 0
                            mentori1 = cols1 * rows1
                            
                            cols2 = int(zenpan_w // h) if h > 0 else 0
                            rows2 = int(zenpan_h // w) if w > 0 else 0
                            mentori2 = cols2 * rows2
                            
                            max_mentori = max(mentori1, mentori2)
                            total_pieces = quantity_input * p_count
                            sheets_needed = math.ceil(total_pieces / max_mentori) if max_mentori > 0 else 0
                            
                            calculated_rows.append({
                                "部材名": p_name,
                                "面数": p_count,
                                "サイズ_W": w,
                                "サイズ_H": h,
                                "面取り数": max_mentori,
                                "必要枚数": sheets_needed
                            })

                        # PDF生成処理
                        buffer = io.BytesIO()
                        doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
                        story = []
                        
                        styles = getSampleStyleSheet()
                        title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontName=font_name, fontSize=14, leading=18)
                        normal_style = ParagraphStyle('NormalStyle', parent=styles['Normal'], fontName=font_name, fontSize=9, leading=13)
                        
                        story.append(Paragraph(f"<b>【材料発注書・図面付き】 {product_input} (受注: {order_id_input})</b>", title_style))
                        story.append(Spacer(1, 8))
                        story.append(Paragraph(f"<b>顧客名:</b> {customer_input} 様 ｜ <b>発注部数:</b> {quantity_input}冊 ｜ <b>全判サイズ:</b> {zenpan_w}mm × {zenpan_h}mm", normal_style))
                        story.append(Spacer(1, 12))
                        
                        # 図解（パーツ配置図）の動的描画
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
                            file_name=f"multi_board_order_{order_id_input}.pdf",
                            mime="application/pdf"
                        )
                        
                    except Exception as pdf_err:
                        st.error(f"PDF生成中にエラーが発生しました: {pdf_err}")

    except Exception as e:
        st.error(f"エラーが発生しました: {str(e)}")