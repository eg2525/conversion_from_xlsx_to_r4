import streamlit as st
import pandas as pd
from io import BytesIO

def app6():
    # タイトル
    st.title("トコロ")

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
        account_options = ["現金", "立替経費", "立替経費（永野）","立替経費（長谷川）", "短期借入金"]

        # ユーザーが選択したオプションをdefault_valueとして直接設定
        selected_default = st.selectbox("科目のデフォルトを選択してください", account_options)
        default_value = selected_default  # 選択された値をそのまま使用
        
        # OKボタンを配置
        if st.button("OK"):
            # OKボタンが押された場合のみ処理を開始
            df_september = dfs[selected_sheet]
            
            # 空の出力用データフレームを作成
            output_columns = [
                "[表題行]", "日付", "伝票番号", "決算整理仕訳", "借方勘定科目", "借方科目コード", "借方補助科目", "借方取引先", "借方取引先コード",
                "借方部門", "借方品目", "借方メモタグ", "借方セグメント1", "借方セグメント2", "借方セグメント3", "借方金額", "借方税区分",
                "借方税額", "貸方勘定科目", "貸方科目コード", "貸方補助科目", "貸方取引先", "貸方取引先コード", "貸方部門", "貸方品目",
                "貸方メモタグ", "貸方セグメント1", "貸方セグメント2", "貸方セグメント3", "貸方金額", "貸方税区分", "貸方税額", "摘要"
                ]
            
            output_df = pd.DataFrame(columns=output_columns)

            # 各処理を実行
            # ① 年・月・日が全て欠けている行を削除
            df_september = df_september.dropna(subset=['年', '月', '日'], how='all')

            # ② 年・月・日をint型に変換
            df_september[['年', '月', '日']] = df_september[['年', '月', '日']].astype(int)

            # ③ 年・月・日をyyyymmdd形式に変換して伝票日付に転記
            df_september['日付'] = (
                df_september['年'].astype(str) + '/' +
                df_september['月'].apply(lambda x: f"{x:02}") + '/' +
                df_september['日'].apply(lambda x: f"{x:02}")
            )
            output_df['日付'] = df_september['日付']

            # ④ 入金・出金の処理
            df_september['借方金額'] = df_september['支払金額']
            df_september['貸方金額'] = df_september['借方金額']
            output_df['借方金額'] = df_september['借方金額'].astype(str)
            output_df['貸方金額'] = df_september['貸方金額'].astype(str)

            # ⑤ 摘要の転記
            df_september['摘要'] = df_september['支払先'] + ' ' + df_september['内容']
            output_df['摘要'] = df_september['摘要']

            # '科目マスタ'シートの'選択肢一覧'列と'勘定科目'列を辞書型にする
            df_master = dfs['科目マスタ']
            account_dict = pd.Series(df_master['勘定科目'].values, index=df_master['選択肢一覧']).to_dict()

            # 借方科目を取得する関数の定義
            def get_debit_account(row):
                # '分類'と照合し、一致するものがあれば科目コードを返し、なければ仮払金を返す
                return account_dict.get(row['分類'], '仮払金')

            # '借方科目'列に結果を出力
            df_september['借方科目'] = df_september.apply(get_debit_account, axis=1)
            output_df['借方勘定科目'] = df_september['借方科目']

            # 貸方科目にデフォルト値を設定（指定されたdefault_valueを使用）
            df_september['貸方科目'] = df_september.get('貸方科目', pd.Series(default_value, index=df_september.index)).fillna(default_value)
            output_df['貸方勘定科目'] = df_september['貸方科目']
                        
            # 借方税区分を条件に基づいて設定
            output_df['借方税区分'] = df_september.apply(
                lambda row: (
                    '不課税' if row['非課税/不課税'] in ['○', '〇'] else
                    '課対仕入（控80）8%（軽）' if row['軽減税率'] in ['○', '〇'] and row['インボイス'] == '登録なし' else
                    '課対仕入8%（軽）' if row['軽減税率'] in ['○', '〇'] and (pd.isna(row['インボイス']) or row['インボイス'] == '') else
                    '課対仕入（控80）10%' if (pd.isna(row['軽減税率']) or row['軽減税率'] == '') and row['インボイス'] == '登録なし' else
                    None
                ),
                axis=1
            )

            # 文字列型が含まれていても対応できるようにfloat型に変換後、整数型に変換
            output_df['借方金額'] = pd.to_numeric(output_df['借方金額'], errors='coerce').fillna(0).astype(int)
            output_df['貸方金額'] = pd.to_numeric(output_df['貸方金額'], errors='coerce').fillna(0).astype(int)

            # CSVファイルをバイナリデータとしてエンコード
            csv_buffer = BytesIO()
            output_df.to_csv(csv_buffer, encoding='cp932', index=False)
            csv_buffer.seek(0)  # バッファの先頭に移動

            # CSVファイルのダウンロードボタン
            st.download_button(label="CSVダウンロード", data=csv_buffer, file_name="output_tkr.csv", mime="text/csv")

            # 完了メッセージ
            st.success("処理が完了しました。CSVファイルをダウンロードできます。")
