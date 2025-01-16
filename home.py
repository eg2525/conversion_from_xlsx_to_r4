import streamlit as st
import template
import excel_to_R4_kaneko
import excel_to_R4
import excel_to_R4_bumon
import excel_to_R4_keihi
import excel_to_R4_yuwa
import excel_to_freee

# 初期状態では何も表示しないようにセッション状態を設定
if 'current_app' not in st.session_state:
    st.session_state['current_app'] = None

st.title('EG_現金出納帳取込アプリ')


# ボタンが押されたときに実行する関数
def show_app0():
    st.session_state['current_app'] = 'app0'

def show_app1():
    st.session_state['current_app'] = 'app1'

def show_app2():
    st.session_state['current_app'] = 'app2'

def show_app3():
    st.session_state['current_app'] = 'app3'

def show_app4():
    st.session_state['current_app'] = 'app4'

def show_app5():
    st.session_state['current_app'] = 'app5'

def show_app6():
    st.session_state['current_app'] = 'app6'

# サイドバーでアプリ選択用のボタンを表示
with st.sidebar:
    if st.button('テンプレート保存', key='app0'):
        show_app0()
    if st.button('金子宝泉堂', key='app1'):
        show_app1()    
    if st.button('部門なし', key='app2'):
        show_app2()    
    if st.button('部門あり', key='app3'):
        show_app3()    
    if st.button('経費取込', key='app4'):
        show_app4()  
    if st.button('結和', key='app5'):
        show_app5()  
    if st.button('トコロ', key='app6'):
        show_app6()  

# 選択されたアプリを表示
if st.session_state['current_app'] == 'app0':
    template.app0()
elif st.session_state['current_app'] == 'app1':
    excel_to_R4_kaneko.app1()
elif st.session_state['current_app'] == 'app2':
    excel_to_R4.app2()
elif st.session_state['current_app'] == 'app3':
    excel_to_R4_bumon.app3()
elif st.session_state['current_app'] == 'app4':
    excel_to_R4_keihi.app4()
elif st.session_state['current_app'] == 'app5':
    excel_to_R4_yuwa.app5()
elif st.session_state['current_app'] == 'app6':
    excel_to_freee.app6()