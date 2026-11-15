import streamlit as st
import pandas as pd
import requests

st.title("🔍 アプリ50 全フィールド値の総当たり調査ツール")

with st.sidebar:
    st.header("検索設定")
    search_val = st.text_input("探したい番号（例: 260001）", value="260001")
    run_diag = st.button("全フィールドから探す")

if run_diag:
    try:
        api_key_50 = st.secrets["app50"]["api_key"]
        url_50 = "https://app050.at-pocket.com/seihon02_bb/api/apps/50/records"
        headers_50 = {"X-At-Pocket-API-Key": api_key_50, "Content-Type": "application/json"}
        
        response = requests.get(url_50, headers=headers_50, timeout=30)
        if response.status_code == 200:
            records = response.json().get("records", [])
            st.success(f"アプリ50から {len(records)} 件のレコードを取得しました。")
            
            matched_records = []
            all_fields_summary = []
            
            for idx, rec in enumerate(records):
                row_data = {"レコード番号(index)": idx}
                is_match = False
                
                for k, v in rec.items():
                    # 値の抽出（辞書ならvalueを取り出す）
                    val = v.get("value", v) if isinstance(v, dict) else v
                    val_str = str(val)
                    row_data[k] = val_str
                    
                    # 探したい値が含まれているかチェック
                    if search_val in val_str:
                        is_match = True
                
                all_fields_summary.append(row_data)
                if is_match:
                    matched_records.append((idx, rec))
            
            if matched_records:
                st.balloons()
                st.success(f"🎉 「{search_val}」が含まれているレコードが {len(matched_records)} 件見つかりました！")
                for idx, mrec in matched_records:
                    st.write(f"### 【レコード #{idx} の中身】")
                    st.json(mrec)
            else:
                st.warning(f"⚠️ どのフィールドにも「{search_val}」という文字は見つかりませんでした。")
                
            st.write("### 📋 全レコードのフィールド値一覧:")
            df_summary = pd.DataFrame(all_fields_summary)
            st.dataframe(df_summary, use_container_width=True)
            
        else:
            st.error(f"API接続エラー: Status {response.status_code}")
    except Exception as e:
            st.error(f"エラー発生: {e}")