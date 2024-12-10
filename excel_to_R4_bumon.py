import streamlit as st
import pandas as pd
from io import BytesIO

def app3():
    # タイトル
    st.title("部門ありExcel")

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
            df_september = dfs[selected_sheet]
            
            # '年', '月', '日'が全て欠けている行を削除
            df_september = df_september.dropna(subset=['年', '月', '日'], how='all')
            
            # '年', '月', '日'を整数型に変換
            df_september[['年', '月', '日']] = df_september[['年', '月', '日']].astype(int)
            
            # 科目マスタの辞書を作成
            df_master = dfs['科目マスタ']
            
            # strip()でデータ整形
            sales_account_dict = {str(k).strip(): v for k, v in zip(df_master['売上科目一覧'], df_master['売上科目コード'])}
            expense_account_dict = {str(k).strip(): v for k, v in zip(df_master['費用科目一覧'], df_master['費用科目コード'])}
            department_dict = {str(k).strip(): v for k, v in zip(df_master['部門一覧'], df_master['部門コード'])}
            
            def get_credit_account(row):
                if pd.notna(row['入金科目']):
                    return sales_account_dict.get(str(row['入金科目']).strip(), default_value)
                else:
                    return default_value

            def get_debit_account(row):
                if pd.notna(row['出金科目']):
                    return expense_account_dict.get(str(row['出金科目']).strip(), default_value)
                else:
                    return default_value
            
            def get_department_code(row):
                if pd.notna(row['部門']):
                    dept_name = str(row['部門']).strip()
                    return department_dict.get(dept_name, 0)
                else:
                    return 0

            # 出力データの列名を定義
            output_columns = [
                "月種別", "種類", "形式", "作成方法", "付箋", "伝票日付", "伝票番号", "伝票摘要", "枝番", 
                "借方部門", "借方部門名", "借方科目", "借方科目名", "借方補助", "借方補助科目名", "借方金額", 
                "借方消費税コード", "借方消費税業種", "借方消費税税率", "借方資金区分", "借方任意項目１", 
                "借方任意項目２", "借方インボイス情報", "貸方部門", "貸方部門名", "貸方科目", "貸方科目名", 
                "貸方補助", "貸方補助科目名", "貸方金額", "貸方消費税コード", "貸方消費税業種", "貸方消費税税率", 
                "貸方資金区分", "貸方任意項目１", "貸方任意項目２", "貸方インボイス情報", "摘要", "期日", "証番号", 
                "入力マシン", "入力ユーザ", "入力アプリ", "入力会社", "入力日付"
            ]
            
            # 出力用エントリリストを初期化
            output_entries = []
            
            # 各行をループ処理
            for index, row in df_september.iterrows():
                # 伝票日付の作成
                date_str = (
                    str(int(row['年'])) +
                    f"{int(row['月']):02}" +
                    f"{int(row['日']):02}"
                )
                denpyou_date = date_str
                summary = row['摘要']
                department_code = get_department_code(row)
                
                # 基本となるエントリを作成（初期は空にしておき、入金・出金で振り分ける）
                base_entry = {
                    "月種別": "",
                    "種類": "",
                    "形式": "",
                    "作成方法": "",
                    "付箋": "",
                    "伝票日付": denpyou_date,
                    "伝票番号": "",
                    "伝票摘要": "",
                    "枝番": "",
                    "借方部門": "",
                    "借方部門名": "",
                    "借方科目": "",
                    "借方科目名": "",
                    "借方補助": 0,
                    "借方補助科目名": "",
                    "借方金額": "",
                    "借方消費税コード": "",
                    "借方消費税業種": "",
                    "借方消費税税率": "",
                    "借方資金区分": "",
                    "借方任意項目１": "",
                    "借方任意項目２": "",
                    "借方インボイス情報": "",
                    "貸方部門": "",
                    "貸方部門名": "",
                    "貸方科目": "",
                    "貸方科目名": "",
                    "貸方補助": 0,
                    "貸方補助科目名": "",
                    "貸方金額": "",
                    "貸方消費税コード": "",
                    "貸方消費税業種": "",
                    "貸方消費税税率": "",
                    "貸方資金区分": "",
                    "貸方任意項目１": "",
                    "貸方任意項目２": "",
                    "貸方インボイス情報": "",
                    "摘要": summary,
                    "期日": "",
                    "証番号": "",
                    "入力マシン": "",
                    "入力ユーザ": "",
                    "入力アプリ": "",
                    "入力会社": "",
                    "入力日付": ""
                }
                
                # '入金'の処理: default_valueが借方科目→借方部門にdepartment_code
                if pd.notna(row['入金']) and row['入金'] != 0:
                    entry = base_entry.copy()
                    amount = row['入金']
                    entry['借方金額'] = str(amount)
                    entry['貸方金額'] = str(amount)
                    entry['借方科目'] = default_value  # 現金科目コード
                    entry['貸方科目'] = get_credit_account(row)
                    
                    entry['借方部門'] = department_code
                    
                    # 軽減税率とインボイス処理
                    if row.get('軽減税率') == '○':
                        entry['貸方消費税コード'] = 2
                        entry['貸方消費税税率'] = 'K8'
                    if row.get('ｲﾝﾎﾞｲｽ') == '○':
                        entry['貸方インボイス情報'] = 8
                        
                    output_entries.append(entry)
                
                # '出金'の処理: default_valueが貸方科目→貸方部門にdepartment_code
                if pd.notna(row['出金']) and row['出金'] != 0:
                    entry = base_entry.copy()
                    amount = row['出金']
                    entry['借方金額'] = str(amount)
                    entry['貸方金額'] = str(amount)
                    entry['借方科目'] = get_debit_account(row)
                    entry['貸方科目'] = default_value  # 現金科目コード
                    
                    entry['貸方部門'] = department_code
                    
                    # 軽減税率とインボイス処理
                    if row.get('軽減税率') == '○':
                        entry['借方消費税コード'] = 32
                        entry['借方消費税税率'] = 'K8'
                    if row.get('ｲﾝﾎﾞｲｽ') == '○':
                        entry['借方インボイス情報'] = 8
                        
                    output_entries.append(entry)
            
            # 出力用DataFrameの作成
            output_df = pd.DataFrame(output_entries, columns=output_columns)
            
            # 部門列をfillna(0)で補完し、int型に変換
            output_df['借方部門'] = output_df['借方部門'].fillna(0).astype(int)
            output_df['貸方部門'] = output_df['貸方部門'].fillna(0).astype(int)
            
            # 借方補助と貸方補助のデフォルト値設定
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
