# -*- coding: utf-8 -*-
import streamlit as st

st.set_page_config(page_title="Ferramentas RP", page_icon="✈️", layout="wide")

st.title("✈️ Ferramentas RP")
st.caption("Escolha a ferramenta no menu à esquerda.")

c1, c2 = st.columns(2)
with c1:
    st.subheader("Conversor de Malha")
    st.write("Converte as malhas da GOL, AZUL, LATAM e as manuais no CSV padrão *Malha RP*.")
    st.caption("Entra: arquivos das cias · Sai: um CSV por período")
with c2:
    st.subheader("Atualizador de OPEX")
    st.write("Atualiza o staff (grupo RAMPA) a partir da folha e a quantidade de voos "
             "a partir da malha. As duas partes são independentes.")
    st.caption("Entra: OPEX + folha e/ou malha · Sai: OPEX atualizado")

st.divider()
st.info("Os arquivos enviados ficam só na memória da sessão — nada é gravado no servidor.")
