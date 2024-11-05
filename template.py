import streamlit as st
import requests
def app0():
    # GitHubの各テンプレートExcelファイルのURL
    TEMPLATE_EXCEL_URLS = {
        "部門なし": "https://github.com/yourusername/yourrepo/raw/main/template1.xlsx",
        "部門あり": "https://github.com/yourusername/yourrepo/raw/main/template2.xlsx",
        "その他": "https://github.com/eg2525/conversion_from_xlsx_to_r4/raw/main/test.xlsx"
    }

    # ページタイトル
    st.title("テンプレートExcelダウンロード")

    # 各テンプレートごとにダウンロードボタンを作成
    for pattern_name, url in TEMPLATE_EXCEL_URLS.items():
        # ダウンロードボタンを表示
        st.write(f"{pattern_name} のテンプレートExcelファイルをダウンロード")
        response = requests.get(url)
        st.download_button(
            label=f"{pattern_name} をダウンロード",
            data=response.content,
            file_name=f"{pattern_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
