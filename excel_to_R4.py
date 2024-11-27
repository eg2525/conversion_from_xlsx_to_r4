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
        
        if st.button("OK"):
            df_september = dfs[selected_sheet]
            
            # Empty list to store output entries
            output_entries = []
            
            # Create sales and expense account dictionaries
            df_master = dfs['科目マスタ']
            sales_account_dict = pd.Series(df_master['売上科目コード'].values, index=df_master['売上科目一覧']).to_dict()
            expense_account_dict = pd.Series(df_master['費用科目コード'].values, index=df_master['費用科目一覧']).to_dict()
            
            def get_credit_account(row):
                if pd.notna(row['入金科目']):
                    return sales_account_dict.get(row['入金科目'], default_value)
                else:
                    return default_value  # Or None?
            
            def get_debit_account(row):
                if pd.notna(row['出金科目']):
                    return expense_account_dict.get(row['出金科目'], default_value)
                else:
                    return default_value  # Or None?
            
            # Drop rows where '年', '月', '日' are all NaN
            df_september = df_september.dropna(subset=['年', '月', '日'], how='all')
            
            # Convert '年', '月', '日' to integers
            df_september[['年', '月', '日']] = df_september[['年', '月', '日']].astype(int)
            
            # Iterate over rows
            for index, row in df_september.iterrows():
                # Process date
                date_str = (
                    str(int(row['年'])) +
                    f"{int(row['月']):02}" +
                    f"{int(row['日']):02}"
                )
                denpyou_date = date_str
                summary = row['摘要']
                
                # Base entry
                base_entry = {
                    "伝票日付": denpyou_date,
                    "摘要": summary,
                    # Initialize other columns as needed
                    # For example, '借方補助': 0, '貸方補助': 0
                    "借方補助": 0,
                    "貸方補助": 0,
                    # Initialize '借方消費税コード', etc., as empty or default
                    "借方消費税コード": '',
                    "借方消費税税率": '',
                    "貸方消費税コード": '',
                    "貸方消費税税率": '',
                    "借方インボイス情報": '',
                    "貸方インボイス情報": '',
                }
                
                # Process '入金'
                if pd.notna(row['入金']) and row['入金'] != 0:
                    entry = base_entry.copy()
                    amount = row['入金']
                    entry['借方金額'] = str(amount)
                    entry['貸方金額'] = str(amount)
                    entry['借方科目'] = default_value  # Cash account code
                    entry['貸方科目'] = get_credit_account(row)
                    
                    # Process '軽減税率' and 'ｲﾝﾎﾞｲｽ'
                    if row.get('軽減税率') == '○':
                        entry['貸方消費税コード'] = 2
                        entry['貸方消費税税率'] = 81
                    if row.get('ｲﾝﾎﾞｲｽ') == '○':
                        entry['貸方インボイス情報'] = 8
                    
                    output_entries.append(entry)
                
                # Process '出金'
                if pd.notna(row['出金']) and row['出金'] != 0:
                    entry = base_entry.copy()
                    amount = row['出金']
                    entry['借方金額'] = str(amount)
                    entry['貸方金額'] = str(amount)
                    entry['借方科目'] = get_debit_account(row)
                    entry['貸方科目'] = default_value  # Cash account code
                    
                    # Process '軽減税率' and 'ｲﾝﾎﾞｲｽ'
                    if row.get('軽減税率') == '○':
                        entry['借方消費税コード'] = 32
                        entry['借方消費税税率'] = 81
                    if row.get('ｲﾝﾎﾞｲｽ') == '○':
                        entry['借方インボイス情報'] = 8
                    
                    output_entries.append(entry)
            
            # Create output DataFrame
            output_df = pd.DataFrame(output_entries, columns=output_columns)
            
            # Fill in other columns with default values if necessary
            output_df['借方補助'] = output_df['借方補助'].fillna(0)
            output_df['貸方補助'] = output_df['貸方補助'].fillna(0)
            
            # CSV export as before
            csv_buffer = BytesIO()
            output_df.to_csv(csv_buffer, encoding='cp932', index=False)
            csv_buffer.seek(0)  # Move to the start of the buffer
            
            st.download_button(label="CSVダウンロード", data=csv_buffer, file_name="output_normal.csv", mime="text/csv")
            
            st.success("処理が完了しました。CSVファイルをダウンロードできます。")
