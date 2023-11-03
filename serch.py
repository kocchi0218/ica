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

        # 40から45列目のデータを取得し、文字列に変換する
        search_data = data.iloc[:, 39:45].applymap(str)

        # チェックボックスでフィルタリングする項目を選択
        st.write("表示から除外する項目を選択してください:")
        excluded_items = st.multiselect("除外する項目", items)

        # 項目リストから除外する項目を削除
        included_items = [item for item in items if item not in excluded_items]

        # 検索ボタン
        if st.button("検索"):
            # 指定された項目でフィルタリングし、A列のデータを抽出
            filtered_indices = search_data.apply(lambda x: any(item in x for item in included_items), axis=1)
            filtered_data = data.loc[filtered_indices]
            
            # 結果を表示
            st.write("検索結果:")
            st.dataframe(filtered_data)
            st.write("A列のデータ:")
            st.dataframe(filtered_data.iloc[:, [0]])

    except Exception as e:
        st.error(f"エラーが発生しました: {e}")
