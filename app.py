import streamlit as st
from generate_offer import grunteco
from io import BytesIO

st.header("Рассчет коммерческого предложения")

tonns = st.number_input('Мощность', min_value=10000, max_value=1000000, value=100000, step=5000)
burt_length = st.selectbox('Длина бурта', [20,30,50], index=2)

density = st.number_input('Плотность', min_value=0.0, max_value=1.0, value=0.6, step=0.05)
weeks = st.selectbox('Недель компостирования', [4,5,6,7], index=1)

burt_wall = st.selectbox('Высота боковых стенок', [0,1,1.5], index=1)
euro = st.number_input('Курс евро', min_value=0.0, max_value=100.0, value=65.0, step=1.0)

offer = st.button('Рассчитать')
if offer:
    file = grunteco(tonns,burt_length,density,weeks,burt_wall, euro)

    binary_output = BytesIO()
    file.save(binary_output)

    st.download_button(label = 'скачать коммерческое предложение',
                    data = binary_output.getvalue(),
                    file_name = 'offer.pptx')
