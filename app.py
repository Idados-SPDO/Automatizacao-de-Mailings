import streamlit as st
import ui 
import pandas as pd

pd.set_option('display.precision', 2)

st.set_page_config(
    page_title="APP Autamatização Mailings",
    layout="wide"
)

def main():
    pages = {
        "Importar planilha base": ui.page_carrega_dado,
        "Preencher Indicador Status Sondagem": ui.page_preenche_indicador,
        "Preencher SACE Status Sondagem": ui.page_preenche_status,
        "Preencher SACE Taxa de Resposta": ui.page_preenche_taxa_resposta,
        "Preencher SACE Sondagens Prioritárias": ui.page_preenche_prioritarias  
    }

    if "page" not in st.session_state:
        st.session_state.update({"page": "Importar planilha de preços"})

    if "data_referencia" not in st.session_state:
        st.session_state.update({"data_referencia":None})

    dataframes = [     
        "Dados_status_EC", "Dados_status_ECE", "Dados_status_ECMA", "Dados_status_ECI"
        "Dados_geral_SCM", "Dados_status_SCM", "Dados_placar_SCM",
        "Dados_geral_SCC", "Dados_status_SCC", "Dados_placar_SCC",
        "Dados_geral_SC", "Dados_status_SC", "Dados_placar_SC",
        "Dados_geral_SSV", "Dados_status_SSV","Dados_placar_SSV"
    ]
    
    for dataframe in dataframes:
        if dataframe not in st.session_state:
            st.session_state.update(
                {
                    dataframe:None,
                }
            )

    with st.sidebar:
        st.title("FGV IBRE - SPDO")
        page = st.radio("Menu", tuple(pages.keys()))
        st.markdown('---')

    pages[page]()

if __name__ == "__main__":
    main()


