import os, hmac
import streamlit as st
st.set_page_config(page_title='Acceso | Revisor de cuentas',layout='centered')
if 'authenticated' not in st.session_state:
    st.session_state.authenticated=False
if not st.session_state.authenticated:
    st.title('Acceso privado · Revisor de cuentas')
    st.info('Solo para despliegues institucionales autorizados. No cargue datos personales reales en servicios de prueba.')
    with st.form('login'):
        username=st.text_input('Usuario')
        password=st.text_input('Contraseña',type='password')
        login=st.form_submit_button('Ingresar')
    if login:
        expected_user=os.environ.get('REVISOR_USER','')
        expected_pass=os.environ.get('REVISOR_PASSWORD','')
        if expected_user and expected_pass and hmac.compare_digest(username,expected_user) and hmac.compare_digest(password,expected_pass):
            st.session_state.authenticated=True
            st.rerun()
        else:
            st.error('Credenciales incorrectas o acceso no configurado.')
    st.stop()
with open(os.path.join(os.path.dirname(__file__),'revisor.py'),encoding='utf-8') as f:
    code=f.read().replace("st.set_page_config(page_title='Revisor de cuentas de cobro',layout='wide')",'')
exec(compile(code,'revisor.py','exec'),{'__name__':'__main__','__file__':__file__})
