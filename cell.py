import streamlit as st
import pandas as pd
import re  # 正規表現を使用するためのモジュール

# ユーザーがファイルをアップロードできるようにする
uploaded_file = st.file_uploader("Excelファイルをアップロードしてください", type=['xlsx'])

if uploaded_file is not None:
    # Excelファイルを読み込む
    df = pd.read_excel(uploaded_file)

    # 指定された列範囲でアスタリスクが付いた項目を検索
    specified_columns = df.iloc[:, 39:45]

    # 結果を格納するためのリストを初期化
    results = []

    # 各列についてアスタリスクを含む項目を検索
    for col_index in range(specified_columns.shape[1]):
        column = specified_columns.columns[col_index]
        for item in specified_columns[column].dropna():
            # アスタリスクを含むか確認
            if '※' in str(item):
                # 数字を抽出
                numbers = re.findall(r'\d+', str(item))
                for number in numbers:
                    # A列で数字を含む項目を検索（列名の代わりにインデックス0を使用）
                    a_column_matches = df.iloc[:, 0].astype(str).str.contains(number, regex=False)
                    if a_column_matches.any():
                        # 該当するA列の情報を取得
                        for a_index in a_column_matches[a_column_matches].index:
                            a_column_value = df.iloc[a_index, 0]
                            results.append((a_column_value, item))

    # 結果をDataFrameに変換
    results_df = pd.DataFrame(results, columns=['A列の情報', '指定列の値'])

    # 結果をStreamlitに表示
    if not results_df.empty:
        st.write("アスタリスク付きの項目と対応するA列の情報:")
        st.dataframe(results_df)
    else:
        st.write("※印のついている項目は見つかりませんでした。")
