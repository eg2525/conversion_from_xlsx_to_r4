import streamlit as st
import pandas as pd
from io import BytesIO

def app2():
    # タイトル
    st.title("部門なしExcel")

    # 1. Excelファイルのアップロード
    uploaded_file = st.file_uploader("Excelファイルをアップロードしてください", type=["xlsx"])

    # ファイルがアップロードされている場合
    if uploaded_file:
        # Excelファイルの全シートを読み込み
        dfs = pd.read_excel(uploaded_file, sheet_name=None)
        
        # シートの存在確認
        if '科目マスタ' not in dfs:
            st.error("アップロードされたファイルに '科目マスタ' シートが含まれていません。")
            return
        
        # 必須データの取得
        df_master = dfs['科目マスタ']
        if '売上科目コード' not in df_master.columns or '売上科目一覧' not in df_master.columns:
            st.error("'科目マスタ' シートに必要な列 ('売上科目コード', '売上科目一覧') がありません。")
            return

        # 科目辞書の作成
        sales_account_dict = pd.Series(
            df_master['売上科目コード'].values, 
            index=df_master['売上科目一覧']
        ).to_dict()
        expense_account_dict = pd.Series(
            df_master['費用科目コード'].values, 
            index=df_master['費用科目一覧']
        ).to_dict()

        # シート選択ドロップダウンを表示
        sheet_names = list(dfs.keys())
        selected_sheet = st.selectbox("シートを選択してください", sheet_names)
        
        # 借方科目と貸方科目の共通デフォルト値を選択肢として表示
        account_options = {
            "現金(100)": 100,
            "立替経費(214)": 214,
            "立替経費(230)": 230
        }
        selected_default = st.selectbox("科目のデフォルトを選択してください", list(account_options.keys()))
        default_value = account_options[selected_default]  # 選択した値を共通デフォルト値として設定
        
        # OKボタンを配置
        if st.button("OK"):
            # OKボタンが押された場合のみ処理を開始
            df_september = dfs[selected_sheet]

            # 空の出力用データフレームを作成
            output_columns = [
                "月種別", "種類", "形式", "作成方法", "付箋", "伝票日付", "伝票番号", "伝票摘要", "枝番", 
                "借方部門", "借方部門名", "借方科目", "借方科目名", "借方補助", "借方補助科目名", "借方金額", 
                "借方消費税コード", "借方消費税業種", "借方消費税税率", "借方資金区分", "借方任意項目１", 
                "借方任意項目２", "借方インボイス情報", "貸方部門", "貸方部門名", "貸方科目", "貸方科目名", 
                "貸方補助", "貸方補助科目名", "貸方金額", "貸方消費税コード", "貸方消費税業種", "貸方消費税税率", 
                "貸方資金区分", "貸方任意項目１", "貸方任意項目２", "貸方インボイス情報", "摘要", "期日", "証番号", 
                "入力マシン", "入力ユーザ", "入力アプリ", "入力会社", "入力日付"
            ]
            output_df = pd.DataFrame(columns=output_columns)

            # 年・月・日が全て欠けている行を削除
            df_september = df_september.dropna(subset=['年', '月', '日'], how='all')

            # 年・月・日をint型に変換
            df_september[['年', '月', '日']] = df_september[['年', '月', '日']].astype(int)

            # 年・月・日をyyyymmdd形式に変換して伝票日付に転記
            df_september['伝票日付'] = (
                df_september['年'].astype(str) +
                df_september['月'].apply(lambda x: f"{x:02}") +
                df_september['日'].apply(lambda x: f"{x:02}")
            )
            output_df['伝票日付'] = df_september['伝票日付']

            # 入出金両方に値がある行を分割
            rows_to_split = df_september[(df_september['入金'].notna()) & (df_september['出金'].notna())]

            # 分割する行の処理
            split_rows = []
            for _, row in rows_to_split.iterrows():
                # 入金用の行
                credit_row = row.copy()
                credit_row['出金'] = None  # 出金をクリア
                split_rows.append(credit_row)

                # 出金用の行
                debit_row = row.copy()
                debit_row['入金'] = None  # 入金をクリア
                split_rows.append(debit_row)

            # 元のデータフレームに分割行を追加し、重複行を削除
            split_df = pd.DataFrame(split_rows)
            df_september = pd.concat([df_september, split_df]).drop_duplicates(keep=False)

            # 借方金額と貸方金額の処理
            df_september['借方金額'] = df_september['出金'].fillna(0)
            df_september['貸方金額'] = df_september['入金'].fillna(0)

            # 借方科目と貸方科目の適用
            def get_credit_account(row):
                return sales_account_dict.get(row['入金科目'], default_value) if pd.notna(row['入金科目']) else None

            def get_debit_account(row):
                return expense_account_dict.get(row['出金科目'], default_value) if pd.notna(row['出金科目']) else None

            df_september['貸方科目'] = df_september.apply(get_credit_account, axis=1)
            df_september['借方科目'] = df_september.apply(get_debit_account, axis=1)
            output_df['貸方科目'] = df_september['貸方科目']
            output_df['借方科目'] = df_september['借方科目']

            # CSV出力
            csv_buffer = BytesIO()
            output_df.to_csv(csv_buffer, encoding='cp932', index=False)
            csv_buffer.seek(0)

            st.download_button(label="CSVダウンロード", data=csv_buffer, file_name="output_normal.csv", mime="text/csv")
            st.success("処理が完了しました。CSVファイルをダウンロードできます。")
