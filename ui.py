import streamlit as st
from openpyxl import load_workbook
import pandas as pd
import data_processing as dp
import datetime as dt
from io import BytesIO

def page_carrega_dado():

    # Cabeçalho da página
    st.header('Ferramenta de :red[Automatização de Mailings]', divider="red")

    # Espaçamento
    st.write('')
    st.write('')
    st.write('')
    st.write('')

    # Crio um container para 'encapsular' itens dentro de uma parte da tela
    container = st.container()

    # Solicito que o usuário selecione uma data e imprimo na tela a data selecionada
    container.subheader('Digite a data de Referência para preenchimento das planilhas:')
    data_referencia = container.date_input("", value=None, format="DD/MM/YYYY")

    #Verifico se o usuário inputou algum valor em data_referencia para poder prosseguir
    if data_referencia is not None:
        container.write(f'Data selecionada: {dp.dia(data_referencia)}/{dp.mes(data_referencia)}/{dp.ano(data_referencia)}')

        # Solicito que o usuário insira as 4 planilhas base para o preenchimento das demais
        container.subheader('Importe os 4 arquivos de relatório utilizados para o preenchimento das planilhas:')
        uploaded_files = container.file_uploader('', accept_multiple_files=True, type=["xlsx"])

        # Crio uma sessão com uma animação de carregamento enquanto salvo os dados na memória temporária da ferramenta para utilizá-los nas páginas de preenchimento subsequentes
        with st.spinner('Preparando e guardando os dados de preenchimento. Por favor aguarde...'):

            # Guardo o valor desta variável na memória temporária da ferramenta para utilizá-la nas páginas de preenchimento subsequentes
            st.session_state.update({"data_referencia":data_referencia})
            
            # Verifico se todos os 4 arquivos foram inseridos pelo usuário. Caso tenham sido, guardo as sheets de interesse dentro da memória do programa também
            if len(uploaded_files) == 4:
                for uploaded_file in uploaded_files:
                    if uploaded_file.name.startswith('SCM_'):
                        st.session_state.update({"Dados_geral_SCM":dp.load_mailing_geral(uploaded_file)}),
                        st.session_state.update({"Dados_status_SCM":dp.load_mailing_status(uploaded_file)}),
                        st.session_state.update({"Dados_placar_SCM":dp.load_mailing_placar(uploaded_file)})
                    if uploaded_file.name.startswith('SCC_'):
                        st.session_state.update({"Dados_geral_SCC":dp.load_mailing_geral(uploaded_file)}),
                        st.session_state.update({"Dados_status_SCC":dp.load_mailing_status(uploaded_file)}),
                        st.session_state.update({"Dados_placar_SCC":dp.load_mailing_placar(uploaded_file)})
                    if uploaded_file.name.startswith('SC_'):
                        st.session_state.update({"Dados_geral_SC":dp.load_mailing_geral(uploaded_file)}),
                        st.session_state.update({"Dados_status_SC":dp.load_mailing_status(uploaded_file)}),
                        st.session_state.update({"Dados_placar_SC":dp.load_mailing_placar(uploaded_file)})
                    if uploaded_file.name.startswith('SSV_'):
                        st.session_state.update({"Dados_geral_SSV":dp.load_mailing_geral(uploaded_file)}),
                        st.session_state.update({"Dados_status_SSV":dp.load_mailing_status(uploaded_file)}),
                        st.session_state.update({"Dados_placar_SSV":dp.load_mailing_placar(uploaded_file)})

                # Gero uma mensagem de sucesso na importação para o usuário
                st.success('Arquivos importados com sucesso!')

        # Espaçamento
        container.write('')
        container.write('')
        container.write('')
        container.write('')

def page_preenche_indicador():

    # Verifico se existem dados para o preenchimento
    if st.session_state.Dados_geral_SCM is not None:

        # Cabeçalho da página
        st.header('Preencher :red[Indicador Status Sondagem]', divider="red")

        # Espaçamento
        st.write('')
        st.write('')
        st.write('')
        st.write('')

        # Crio um container para 'encapsular' itens dentro de uma parte da tela
        container = st.container()
        container.subheader('Importe arquivo a ser atualizado:')

        # Lista onde vou guardar o arquivo a ser modificado pelo usuário
        uploaded_indicador = container.file_uploader('', accept_multiple_files=False, type=["xlsx"], key="page1")

        # Verifico se a lista tem algum arquivo
        if uploaded_indicador is not None:

            # Lista com as sondagens
            sondagens = ['Comércio', 'Construção', 'Indústria', 'Serviços'] # Utilizo para definir o nome da sheet a ser lida no momento da modificação

            # Lista com os dfs utilizados para modificar o arquivo inserido
            dfs_status = [
                st.session_state.Dados_status_SCM, 
                st.session_state.Dados_status_SCC, 
                st.session_state.Dados_status_SC, 
                st.session_state.Dados_status_SSV
            ] # Utilizo para definir o arquivo utilizado na modificação de cada sheet

            # Carrego a planilha existente
            workbook = load_workbook(uploaded_indicador)

            # Loop para fazer as modificações de forma recursiva para cada sondagem
            for i in range(len(sondagens)):

                # Carregar a sheet da planilha existente                
                sheet_workbook = workbook[sondagens[i]]

                # Modifico o valor de df_status de acordo com a sondagem que será modificada
                df_status = dfs_status[i]
                
                # Redefino o valor de workbook, que inicialmente era igual ao arquivo passado pelo usuário, para o arquivo já modificado pela função preencher_indicador
                workbook = dp.preenche_indicador(df_status, workbook, sondagens[i])

            # Salvo o workbook em um buffer de memória
            buffer = BytesIO()
            workbook.save(buffer)
            buffer.seek(0)

            # Espaçamento
            st.write('')
            st.write('')
            st.write('')
            st.write('')

            # Permito o download do novo arquivo atualizado
            st.download_button(
                label="Baixar arquivo atualizado",
                data=buffer,
                file_name=f'Indicador_STATUS_Sondagem - Atualizado_{dp.dia(dt.date.today())}/{dp.mes(dt.date.today())}.xlsx',
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # Caso não existam dados para o preenchimento, retorno uma mensagem de aviso para o usuário
    else:
        st.warning('Por favor insira as Planilhas em "Importar planilha base".')

def page_preenche_status():   
        
    # Verifico se existem dados para o preenchimento
    if st.session_state.Dados_geral_SCM is not None:

        # Cabeçalho da página
        st.header('Preencher :red[SACE Status Sondagem]', divider="red")

        # Espaçamento
        st.write('')
        st.write('')
        st.write('')
        st.write('')

        container = st.container()
        container.subheader('Importe arquivo a ser atualizado:')

        # Lista onde vou guardar o arquivo a ser modificado pelo usuário
        uploaded_status = container.file_uploader('', accept_multiple_files=False, type=["xlsx"], key="page2")

        # Verifico se a lista tem algum arquivo
        if uploaded_status is not None:

            # Lista com as sondagens
            sondagens = ['Comércio', 'Construção', 'Indústria', 'Serviços'] # Utilizo para definir o nome da sheet a ser lida no momento da modificação

            # Lista com os dfs utilizados para modificar o arquivo inserido
            dfs_status = [
                st.session_state.Dados_status_SCM, 
                st.session_state.Dados_status_SCC, 
                st.session_state.Dados_status_SC, 
                st.session_state.Dados_status_SSV
            ] # Utilizo para definir o arquivo utilizado na modificação de cada sheet

            # Carrego a planilha existente
            workbook = load_workbook(uploaded_status)

            # Loop para fazer as modificações de forma recursiva para cada sondagem
            for i in range(len(sondagens)):

                # Modifico o valor de df_status de acordo com a sondagem que será modificada
                df_status = dfs_status[i]
                
                # Redefino o valor de workbook, que inicialmente era igual ao arquivo inicial, para o arquivo já modificado pela função preencher_status
                workbook = dp.preenche_status(df_status, workbook, sondagens[i].upper()) 

            # Salvo o workbook em um buffer de memória
            buffer = BytesIO()
            workbook.save(buffer)
            buffer.seek(0)

            # Espaçamento
            st.write('')
            st.write('')
            st.write('')
            st.write('')

            # Permito o download do novo arquivo atualizado
            st.download_button(
                label="Baixar arquivo atualizado",
                data=buffer,
                file_name= f'BD_SACE_Sondagem_Status_{dp.dia(dt.date.today())}.{dp.mes(dt.date.today())}.xlsx',
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # Caso não existam dados para o preenchimento, retorno uma mensagem de aviso para o usuário
    else:
        st.warning('Por favor insira as Planilhas em "Importar planilha base".')

def page_preenche_taxa_resposta():

     # Verifico se existem dados para o preenchimento
    if st.session_state.Dados_geral_SCM is not None:

        # Cabeçalho da página
        st.header('Preencher :red[SACE Sondagem Taxa de Resposta]', divider="red")

        # Espaçamento
        st.write('')
        st.write('')
        st.write('')
        st.write('')

        container = st.container()
        container.subheader('Importe arquivo a ser atualizado:')

        # Lista onde vou guardar o arquivo a ser modificado pelo usuário
        uploaded_taxa = container.file_uploader('', accept_multiple_files=False, type=["xlsx"], key="page3")

        # Verifico se a lista tem algum arquivo
        if uploaded_taxa is not None:

            # Lista com as sondagens
            sondagens = ['Comércio', 'Construção', 'Indústria', 'Serviços'] # Utilizo para definir o nome da sheet a ser lida no momento da modificação

            # Lista com os dfs utilizados para modificar o arquivo inserido
            dfs_taxa = [
                st.session_state.Dados_geral_SCM, 
                st.session_state.Dados_geral_SCC, 
                st.session_state.Dados_geral_SC, 
                st.session_state.Dados_geral_SSV
            ] # Utilizo para definir o arquivo utilizado na modificação de cada sheet

            # Guardo o arquivo passado pelo usuário em formato de df
            df_uploaded = pd.read_excel(uploaded_taxa)

            # Loop para fazer as modificações de forma recursiva para cada sondagem
            for i in range(len(sondagens)):

                # Modifico o valor de df_taxa de acordo com a sondagem que será modificada
                df_taxa = dfs_taxa[i]
                
                # Crio df_atualizado para guardar as informações de df_uploaded modificado pela função preenche_taxa_resposta
                df_atualizado = dp.preenche_taxa_resposta(df_taxa, df_uploaded, sondagens[i]) 

                # Redefino o valor de df_uploaded para que na próxima rodagem de cada iteração ele receba as modificações da iteração atual
                df_uploaded = df_atualizado

            # Salvo o df_atualizado em um buffer de memória
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df_atualizado.to_excel(writer, index = False)
            buffer.seek(0)

            # Dou ao usuário uma prévia do conteúdo do arquivo que será baixado por ele
            st.dataframe(pd.read_excel(buffer), width=800, height=300)

            # Espaçamento
            st.write('')
            st.write('')
            st.write('')
            st.write('')

            # Permito o download do novo arquivo atualizado
            st.download_button(
                label="Baixar arquivo atualizado",
                data=buffer,
                file_name= f'BD_SACE_Sondagem_Taxa de Respostas_{dp.dia(dt.date.today())}.{dp.mes(dt.date.today())}.xlsx',
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # Caso não existam dados para o preenchimento, retorno uma mensagem de aviso para o usuário
    else:
        st.warning('Por favor insira as Planilhas em "Importar planilha base".')

def page_preenche_prioritarias(): 
        
    # Verifico se existem dados para o preenchimento
    if st.session_state.Dados_geral_SCM is not None:

        # Cabeçalho da página
        st.header('Preencher :red[SACE Sondagens Prioritárias]', divider="red")

        # Espaçamento
        st.write('')
        st.write('')
        st.write('')
        st.write('')

        container = st.container()
        container.subheader('Importe arquivo a ser atualizado:')

        # Lista onde vou guardar o arquivo a ser modificado pelo usuário
        uploaded_prioritarias = container.file_uploader('', accept_multiple_files=False, type=["xlsx"], key="page4")

        # Verifico se a lista tem algum arquivo
        if uploaded_prioritarias is not None:

            # Lista com as sondagens
            sondagens = ['Comércio', 'Construção', 'Indústria', 'Serviços'] # Utilizo para definir o nome da sheet a ser lida no momento da modiMicrosoft Teamsficação

            # Lista com os dfs utilizados para modificar o arquivo inserido
            dfs_placar = [
                st.session_state.Dados_placar_SCM, 
                st.session_state.Dados_placar_SCC, 
                st.session_state.Dados_placar_SC, 
                st.session_state.Dados_placar_SSV
            ] # Utilizo para definir o arquivo utilizado na modificação de cada sheet

            # Crio uma sessão com uma animação de carregamento enquanto as modificações são geradas
            with st.spinner('Preenchendo a nova planilha. Por favor aguarde...'):
                
                # Carrego a planilha existente
                workbook = load_workbook(uploaded_prioritarias)

           
                # Loop para fazer as modificações de forma recursiva para cada sondagem
                for i in range(len(sondagens)):

                    # Modifico o valor de df_status de acordo com a sondagem que será modificada
                    df_placar = dfs_placar[i]

                    # Redefino o valor de workbook, que inicialmente era igual ao arquivo inicial, para o arquivo já modificado pela função preencher_status
                    workbook = dp.preenche_prioritarias(df_placar, workbook, sondagens[i]) 

                # Salvo o workbook em um buffer de memória
                buffer = BytesIO()
                workbook.save(buffer)
                buffer.seek(0)

            # Espaçamento
            st.write('')
            st.write('')
            st.write('')
            st.write('')

            # Permito o download do novo arquivo atualizado
            st.download_button(
                label="Baixar arquivo atualizado",
                data=buffer,
                file_name= f'BD_SACE_Sondagem_Prioritárias_{dp.dia(dt.date.today())}.{dp.mes(dt.date.today())}.xlsx',
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # Caso não existam dados para o preenchimento, retorno uma mensagem de aviso para o usuário
    else:
        st.warning('Por favor insira as Planilhas em "Importar planilha base".')