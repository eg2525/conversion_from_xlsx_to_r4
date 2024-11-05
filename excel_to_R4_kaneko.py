import streamlit as st
import pandas as pd
from io import BytesIO
def app1():
    # タイトル
    st.title("金子宝泉堂")

    # 1. Excelファイルのアップロード
    uploaded_file = st.file_uploader("Excelファイルをアップロードしてください", type=["xlsx"])

    # ファイルがアップロードされている場合
    if uploaded_file:
        # Excelファイルの全シートを読み込み
        dfs = pd.read_excel(uploaded_file, sheet_name=None)
        
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

            # 各処理を実行
            # ① 年・月・日が全て欠けている行を削除
            df_september = df_september.dropna(subset=['年', '月', '日'], how='all')

            # ② 年・月・日をint型に変換
            df_september[['年', '月', '日']] = df_september[['年', '月', '日']].astype(int)

            # ③ 年・月・日をyyyymmdd形式に変換して伝票日付に転記
            df_september['伝票日付'] = (
                df_september['年'].astype(str) +
                df_september['月'].apply(lambda x: f"{x:02}") +
                df_september['日'].apply(lambda x: f"{x:02}")
            )
            output_df['伝票日付'] = df_september['伝票日付']

            # ④ 入金・出金の処理
            df_september['借方金額'] = df_september[['入金', '出金']].sum(axis=1, skipna=True)
            df_september['貸方金額'] = df_september['借方金額']
            output_df['借方金額'] = df_september['借方金額'].astype(int)
            output_df['貸方金額'] = df_september['貸方金額'].astype(int)

            # ⑤ 摘要の転記
            output_df['摘要'] = df_september['摘要']

            # ⑥ '入金科目'と'売上科目一覧'の照合
            df_master = dfs['科目マスタ']
            sales_account_dict = pd.Series(df_master['売上科目コード'].values, index=df_master['売上科目一覧']).to_dict()

            def get_credit_account(row):
                if pd.notna(row['入金科目']):
                    return sales_account_dict.get(row['入金科目'], default_value)
                else:
                    return None

            df_september['貸方科目'] = df_september.apply(get_credit_account, axis=1)
            output_df['貸方科目'] = df_september['貸方科目'].fillna(default_value)

            # ⑦ '出金科目'と'費用科目一覧'の照合
            expense_account_dict = pd.Series(df_master['費用科目コード'].values, index=df_master['費用科目一覧']).to_dict()

            def get_debit_account(row):
                if pd.notna(row['出金科目']):
                    return expense_account_dict.get(row['出金科目'], default_value)
                else:
                    return None

            df_september['借方科目'] = df_september.apply(get_debit_account, axis=1)
            output_df['借方科目'] = df_september['借方科目'].fillna(default_value)

            # ⑧ '軽減税率'確認
            df_september['借方消費税コード'] = df_september['軽減税率'].apply(lambda x: 32 if x == '○' else None)
            df_september['借方消費税税率'] = df_september['軽減税率'].apply(lambda x: 81 if x == '○' else None)
            output_df['借方消費税コード'] = df_september['借方消費税コード']
            output_df['借方消費税税率'] = df_september['借方消費税税率']

            # ⑨ 'ｲﾝﾎﾞｲｽ'確認
            df_september['借方インボイス情報'] = df_september['ｲﾝﾎﾞｲｽ'].apply(lambda x: 8 if x == '○' else None)
            output_df['借方インボイス情報'] = df_september['借方インボイス情報']

            # ⑩ '本部経費'確認
            df_september['借方部門'] = df_september['本部経費'].apply(lambda x: 99 if x == '○' else None)
            output_df['借方部門'] = df_september['借方部門']

            # ⑫ 借方補助と貸方補助のデフォルト値設定
            output_df['借方補助'] = output_df['借方補助'].fillna(0)
            output_df['貸方補助'] = output_df['貸方補助'].fillna(0)

            # CSVファイルをバイナリデータとしてエンコード
            csv_buffer = BytesIO()
            output_df.to_csv(csv_buffer, encoding='cp932', index=False)
            csv_buffer.seek(0)  # バッファの先頭に移動

            # CSVファイルのダウンロードボタン
            st.download_button(label="CSVダウンロード", data=csv_buffer, file_name="output.csv", mime="text/csv")

            # 完了メッセージ
            st.success("処理が完了しました。CSVファイルをダウンロードできます。")
