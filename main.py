import streamlit as st
import pandas as pd

# アプリのタイトル
st.title("Excelデータ検索アプリ")

# ファイルアップローダーを作成
uploaded_file = st.file_uploader("Excelファイルをドラッグ＆ドロップ、またはブラウズしてアップロードしてください", type=['xlsx', 'xls'])

# ファイルがアップロードされたら処理を開始
if uploaded_file is not None:
    try:
        # ファイルを読み込む
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names

        # シート名を選択させる
        sheet_name = st.selectbox("シートを選択してください", options=sheet_names)

        # 選択したシートを読み込む
        data = pd.read_excel(uploaded_file, sheet_name=sheet_name)

        # 検索条件の入力
        search_query = st.text_input("検索条件を入力してください")

        # 検索ボタン
        if st.button("検索"):
            # 検索結果のデータフレームをフィルタリング
            filtered_data = data[data.apply(lambda row: row.astype(str).str.contains(search_query, na=False).any(), axis=1)]

            # 結果を表示
            st.dataframe(filtered_data)

    except Exception as e:
        st.error(f"エラーが発生しました: {e}")
