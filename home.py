import streamlit as st
import excel_to_R4_kaneko


# 初期状態では何も表示しないようにセッション状態を設定
if 'current_app' not in st.session_state:
    st.session_state['current_app'] = None

st.title('EG_現金出納帳取込アプリ')


# ボタンが押されたときに実行する関数
def show_app1():
    st.session_state['current_app'] = 'app1'


# サイドバーでアプリ選択用のボタンを表示
with st.sidebar:
    if st.button('金子宝泉堂', key='app1'):
        show_app1()

# 選択されたアプリを表示
if st.session_state['current_app'] == 'app1':
    excel_to_R4_kaneko.app1()