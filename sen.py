import streamlit as st
import pandas as pd

# アプリのタイトル
st.title("Excelデータ検索アプリ")

# ファイルアップローダーを作成
uploaded_file = st.file_uploader("Excelファイルをドラッグ＆ドロップ、またはブラウズしてアップロードしてください", type=['xlsx', 'xls'])

# 項目の選択肢
items = [
    "ALCP2MP調", "C3DF2調", "C3IC調製品", "C3M7調製品", "GED2調", "VACL2調",
    "YLF4調", "YLZ調", "TEDT調", "TC3ED調", "PCR2調", "CYLACL調", "C3ED調", 
    "C3A7調", "ALJ調", "ALCP4調"
]

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

        # 項目のセレクトボックスを作成
        selected_item = st.selectbox("検索する項目を選択してください", options=items)

        # 検索ボタン
        if st.button("検索"):
            # 40から45列目のデータを取得し、文字列に変換する
            search_data = data.iloc[:, 39:45].applymap(str)
            # 指定された項目でフィルタリングし、A列のデータを抽出
            filtered_indices = search_data.apply(lambda x: x.str.contains(selected_item, regex=False, na=False)).any(axis=1)
            filtered_data = data.loc[filtered_indices]
            a_column_data = filtered_data[data.columns[0]]  # A列のデータを取得

            # 結果を表示
            st.write("検索結果:")
            st.dataframe(filtered_data)
            st.write("A列のデータ:")
            st.dataframe(a_column_data)

            # 選択された項目が含まれるセルの内容を表示
            st.write("選択された項目が含まれるセルの内容:")
            for index, row in filtered_data.iterrows():
                for col in row.index[39:45]:
                    if selected_item in str(row[col]):
                        st.write(f"行 {index + 1}, 列 {col}: {row[col]}")

    except Exception as e:
        st.error(f"エラーが発生しました: {e}")
