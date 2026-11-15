import streamlit as st
import pandas as pd
import requests

st.title("🔍 アプリ50・field-55 診断ツール")

# サイドバーでAPIキー設定
with st.sidebar:
    st.header("API設定確認")
    target_id = st.text_input("探したい受注番号", value="260002")
    check_btn = st.button("field-55 を検索する")

if check_btn:
    try:
        api_key_50 = st.secrets["app50"]["api_key"]
        url_50 = "https://app050.at-pocket.com/seihon02_bb/api/apps/50/records"
        headers_50 = {"X-At-Pocket-API-Key": api_key_50, "Content-Type": "application/json"}
        
        response = requests.get(url_50, headers=headers_50, timeout=30)
        if response.status_code == 200:
            records = response.json().get("records", [])
            st.success(f"アプリ50から {len(records)} 件のレコードを取得しました。")
            
            # field-55 の一覧を調査
            found_rec = None
            all_f55_list = []
            
            for rec in records:
                # field-55 の値を取り出す
                f55_field = rec.get("field-55", "")
                f55_val = str(f55_field.get("value", f55_field) if isinstance(f55_field, dict) else f55_field).strip()
                all_f55_list.append(f55_val)
                
                if target_id in f55_val:
                    found_rec = rec
                    break
            
            if found_rec:
                st.balloons()
                st.success(f"🎉 該当レコードを発見しました！ (field-55の値: {target_id})")
                st.json(found_rec)
            else:
                st.error(f"❌ field-55 の中に「{target_id}」に一致するデータが見つかりませんでした。")
                st.write("### 📋 現在アプリ50に登録されている field-55 の値一覧:")
                st.write(all_f55_list)
        else:
            st.error(f"API接続エラー: Status {response.status_code}")
    except Exception as e:
        st.error(f"エラー発生: {e}")