import streamlit as st
import pandas as pd
from io import BytesIO

def app5():
    # タイトル
    st.title("結和取込")

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
            "立替経費(1193)": 1193,
            "短期借入金(202)": 202
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
            df_september['借方金額'] = df_september['支払金額']
            df_september['貸方金額'] = df_september['借方金額']
            output_df['借方金額'] = df_september['借方金額'].astype(str)
            output_df['貸方金額'] = df_september['貸方金額'].astype(str)

            # ⑤ 摘要の転記
            df_september['摘要'] = df_september['支払先'] + ' ' + df_september['内容']
            output_df['摘要'] = df_september['摘要']

            # '科目マスタ'シートの'選択肢一覧'列と'科目コード'列を辞書型にする
            df_master = dfs['科目マスタ']
            account_dict = pd.Series(df_master['科目コード'].values, index=df_master['選択肢一覧']).to_dict()

            # 借方科目を取得する関数の定義
            def get_debit_account(row):
                # '分類'と照合し、一致するものがあれば科目コードを返し、なければ154を返す
                return account_dict.get(row['分類'], 154)

            # '借方科目'列に結果を出力
            df_september['借方科目'] = df_september.apply(get_debit_account, axis=1)
            output_df['借方科目'] = df_september['借方科目']

            # 貸方科目にデフォルト値を設定（指定されたdefault_valueを使用）
            df_september['貸方科目'] = df_september.get('貸方科目', pd.Series(default_value, index=df_september.index)).fillna(default_value)
            output_df['貸方科目'] = df_september['貸方科目']

            # ⑧ '軽減税率'確認
            df_september['借方消費税コード'] = df_september['軽減税率'].apply(lambda x: 42 if x in ['○', '〇'] else None)
            df_september['借方消費税税率'] = df_september['軽減税率'].apply(lambda x: 'K8' if x in ['○', '〇'] else None)
            output_df['借方消費税コード'] = df_september['借方消費税コード']
            output_df['借方消費税税率'] = df_september['借方消費税税率']

            # ⑨ 'ｲﾝﾎﾞｲｽ'確認
            df_september['借方インボイス情報'] = df_september['インボイス'].apply(lambda x: 8 if x == '登録なし' else None)
            output_df['借方インボイス情報'] = df_september['借方インボイス情報']

            # ⑫ 借方補助と貸方補助のデフォルト値設定
            # 借方科目が527で借方補助がNaNの場合は2を埋める
            output_df.loc[(output_df['借方科目'] == 527) & (output_df['借方補助'].isna()), '借方補助'] = 2
            output_df.loc[(output_df['貸方科目'] == 202) & (output_df['貸方補助'].isna()), '借方補助'] = 1


            # それ以外のNaNは0を埋める
            output_df['借方補助'] = output_df['借方補助'].fillna(0)
            output_df['貸方補助'] = output_df['貸方補助'].fillna(0)
            output_df['借方部門'] = output_df['借方部門'].fillna(1)

            # CSVファイルをバイナリデータとしてエンコード
            csv_buffer = BytesIO()
            output_df.to_csv(csv_buffer, encoding='cp932', index=False)
            csv_buffer.seek(0)  # バッファの先頭に移動

            # CSVファイルのダウンロードボタン
            st.download_button(label="CSVダウンロード", data=csv_buffer, file_name="output_yuwa.csv", mime="text/csv")

            # 完了メッセージ
            st.success("処理が完了しました。CSVファイルをダウンロードできます。")
