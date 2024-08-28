import streamlit as st
from openpyxl import load_workbook
import pandas as pd
import data_processing as dp
import datetime as dt
from io import BytesIO

def page_carrega_dado():

    # Cabeçalho da página.
    st.header('Ferramenta de :red[Automatização de Mailings]', divider="red")

    # Espaçamento.
    st.write('')
    st.write('')
    st.write('')
    st.write('')

    # Crio um container para 'encapsular' itens dentro de uma parte da tela.
    container = st.container()

    # Solicito que o usuário selecione uma data e imprimo na tela a data selecionada.
    container.subheader('Digite a data de Referência para preenchimento das planilhas:')
    data_referencia = container.date_input("", value=None, format="DD/MM/YYYY")

    #Verifico se o usuário inputou algum valor em data_referencia para poder prosseguir.
    if data_referencia is not None:
        container.write(f'Data selecionada: {dp.dia(data_referencia)}/{dp.mes(data_referencia)}/{dp.ano(data_referencia)}')

        # Solicito que o usuário insira pelo menos uma planilha base para o preenchimento das demais.
        container.subheader('Importe os arquivos de relatório utilizados para o preenchimento das planilhas:')
        uploaded_files = container.file_uploader('', accept_multiple_files=True, type=["xlsx"])

        # Crio uma sessão com uma animação de carregamento enquanto salvo os dados na memória temporária da ferramenta para utilizá-los nas páginas de preenchimento subsequentes.
        with st.spinner('Preparando e guardando os dados de preenchimento. Por favor aguarde...'):

            # Guardo o valor desta variável na memória temporária da ferramenta para utilizá-la nas páginas de preenchimento subsequentes.
            st.session_state.update({"data_referencia":data_referencia})
            
            # Verifico se pelo menos um dos mailings foi inserido pelo usuário. Caso tenham sido, guardo a(s) sheet(s) de interesse dentro da memória do programa também.
            if len(uploaded_files) >= 1:
                for uploaded_file in uploaded_files:
                    if uploaded_file.name.startswith('ECE_'):
                        st.session_state.update({"Dados_status_ECE":dp.load_mailing_status(uploaded_file)})
                    if uploaded_file.name.startswith('ECMA_'):
                        st.session_state.update({"Dados_status_ECMA":dp.load_mailing_status(uploaded_file)})
                    if uploaded_file.name.startswith('ECI_'):
                        st.session_state.update({"Dados_status_ECI":dp.load_mailing_status(uploaded_file)})
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

                # Gero uma mensagem de sucesso na importação para o usuário.
                st.success('Arquivos importados com sucesso!')

        # Espaçamento
        container.write('')
        container.write('')
        container.write('')
        container.write('')

def page_preenche_indicador():

    # Crio uma lista com todos os arquivos que poderão ser usados na página.
    arquivos_previstos = [
        st.session_state.Dados_status_ECE,
        st.session_state.Dados_status_ECMA, st.session_state.Dados_status_ECI,
        st.session_state.Dados_status_SCM, st.session_state.Dados_status_SCC,
        st.session_state.Dados_status_SC, st.session_state.Dados_status_SSV
    ]

    arquivos_inseridos = [] # Lista contendo somente os arquivos que de fato foram inseridos pelo usuário.
    sondagens_trabalhadas = [] # Lista contendo as sondagens correspondentes a cada arquivo inserido.
    contador = 0

    # Loop responsável por percorrer todos os arquivos da lista arquivos_previstos e adicionar os que não estão em branco na lista de arquivos_inseridos.
    for arquivo in arquivos_previstos:
        
        if arquivo is not None:
            arquivos_inseridos.append(arquivo)

            # Com base no contador, identifico a sondagem correspondente ao arquivo que foi inserido e a adiciono na lista de sondagens_trabalhadas.
            if contador == 0:
                sondagens_trabalhadas.append('Consumidor (E)')
            if contador == 1:
                sondagens_trabalhadas.append('Consumidor (MA)')
            if contador == 2:
                sondagens_trabalhadas.append('Consumidor (I)')
            if contador == 3:
                sondagens_trabalhadas.append('Comércio')
            if contador == 4:
                sondagens_trabalhadas.append('Construção')
            if contador == 5:
                sondagens_trabalhadas.append('Indústria')
            if contador == 6:
                sondagens_trabalhadas.append('Serviços')

        contador += 1

    # Verifico se existem dados para o preenchimento.
    if arquivos_inseridos:

        # Cabeçalho da página.
        st.header('Preencher :red[Indicador Status Sondagem]', divider="red")

        # Espaçamento.
        st.write('')
        st.write('')
        st.write('')
        st.write('')

        # Crio um container para 'encapsular' itens dentro de uma parte da tela.
        container = st.container()
        container.subheader('Importe arquivo a ser atualizado:')

        # Lista onde vou guardar o arquivo a ser modificado pelo usuário.
        uploaded_indicador = container.file_uploader('', accept_multiple_files=False, type=["xlsx"], key="page1")

        # Verifico se a lista tem algum arquivo.
        if uploaded_indicador is not None:

            # Carrego a planilha existente.
            workbook = load_workbook(uploaded_indicador)

            # Loop para fazer as modificações de forma recursiva para cada sondagem inserida.
            for i in range(len(sondagens_trabalhadas)):

                # Carregar a sheet da planilha existente.                
                sheet_workbook = workbook[sondagens_trabalhadas[i]]

                # Modifico o valor de df_status de acordo com a sondagem que será modificada.
                df_status = arquivos_inseridos[i]
                
                # Redefino o valor de workbook, que inicialmente era igual ao arquivo passado pelo usuário, para o arquivo já modificado pela função preencher_indicador.
                workbook = dp.preenche_indicador(df_status, workbook, sondagens_trabalhadas[i])

            # Salvo o workbook em um buffer de memória.
            buffer = BytesIO()
            workbook.save(buffer)
            buffer.seek(0)

            # Espaçamento.
            st.write('')
            st.write('')
            st.write('')
            st.write('')

            # Permito o download do novo arquivo atualizado.
            st.download_button(
                label="Baixar arquivo atualizado",
                data=buffer,
                file_name=f'Indicador_STATUS_Sondagem - Atualizado_{dp.dia(dt.date.today())}/{dp.mes(dt.date.today())}.xlsx',
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # Caso não existam dados para o preenchimento, retorno uma mensagem de aviso para o usuário.
    else:
        st.warning('Por favor insira as Planilhas em "Importar planilha base".')

def page_preenche_status():   

    # Caso Dado_status_EC, que é a consolidação das 3 subdivisões do Consumidor (ECE, ECMA e ECI), estiver vazio, faço o preencimento deste.
    if st.session_state.Dados_status_EC is None:

        # Guardo os valores de Consumidor salvos em session_state em 3 dataframes que serão utilizados na criação do consolidado.
        df_ECE = st.session_state.Dados_status_ECE
        df_ECMA = st.session_state.Dados_status_ECMA
        df_ECI = st.session_state.Dados_status_ECI

        # Realizo a soma das colunas QUANTIDADE.
        df_EC = df_ECE.copy()
        df_EC['QUANTIDADE'] += df_ECMA['QUANTIDADE'] + df_ECI['QUANTIDADE']

        # Atualizo o Dados_status_EC.
        st.session_state.update({"Dados_status_EC": df_EC}) 

    # Crio uma lista com todos os arquivos que poderão ser usados na página.
    arquivos_previstos = [
        st.session_state.Dados_status_EC,
        st.session_state.Dados_status_SCM, st.session_state.Dados_status_SCC,
        st.session_state.Dados_status_SC, st.session_state.Dados_status_SSV
    ]

    arquivos_inseridos = [] # Lista contendo somente os arquivos que de fato foram inseridos pelo usuário.
    sondagens_trabalhadas = [] # Lista contendo as sondagens correspondentes a cada arquivo inserido.
    contador = 0

    # Loop responsável por percorrer todos os arquivos da lista arquivos_previstos e adicionar os que não estão em branco na lista de arquivos_inseridos.
    for arquivo in arquivos_previstos:
        
        if arquivo is not None:
            arquivos_inseridos.append(arquivo)

            # Com base no contador, identifico a sondagem correspondente ao arquivo que foi inserido e a adiciono na lista de sondagens_trabalhadas.
            if contador == 0:
                sondagens_trabalhadas.append('Consumidor')
            if contador == 1:
                sondagens_trabalhadas.append('Comércio')
            if contador == 2:
                sondagens_trabalhadas.append('Construção')
            if contador == 3:
                sondagens_trabalhadas.append('Indústria')
            if contador == 4:
                sondagens_trabalhadas.append('Serviços')

        contador += 1

    # Verifico se existem dados para o preenchimento.
    if arquivos_inseridos:

        # Cabeçalho da página.
        st.header('Preencher :red[SACE Status Sondagem]', divider="red")

        # Espaçamento.
        st.write('')
        st.write('')
        st.write('')
        st.write('')

        container = st.container()
        container.subheader('Importe arquivo a ser atualizado:')

        # Lista onde vou guardar o arquivo a ser modificado pelo usuário.
        uploaded_status = container.file_uploader('', accept_multiple_files=False, type=["xlsx"], key="page2")

        # Verifico se a lista tem algum arquivo.
        if uploaded_status is not None:

            # Carrego a planilha existente.
            workbook = load_workbook(uploaded_status)

            # Loop para fazer as modificações de forma recursiva para cada sondagem inserida.
            for i in range(len(sondagens_trabalhadas)):

                # Modifico o valor de df_status de acordo com a sondagem que será modificada.
                df_status = arquivos_inseridos[i]
                
                # Redefino o valor de workbook, que inicialmente era igual ao arquivo inicial, para o arquivo já modificado pela função preencher_status.
                workbook = dp.preenche_status(df_status, workbook, sondagens_trabalhadas[i].upper()) 

            # Salvo o workbook em um buffer de memória.
            buffer = BytesIO()
            workbook.save(buffer)
            buffer.seek(0)

            # Dou ao usuário uma prévia do conteúdo do arquivo que será baixado por ele.
            st.dataframe(pd.read_excel(buffer), width=800, height=300)

            # Espaçamento.
            st.write('')
            st.write('')
            st.write('')
            st.write('')

            # Permito o download do novo arquivo atualizado.
            st.download_button(
                label="Baixar arquivo atualizado",
                data=buffer,
                file_name= f'BD_SACE_Sondagem_Status_{dp.dia(dt.date.today())}.{dp.mes(dt.date.today())}.xlsx',
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # Caso não existam dados para o preenchimento, retorno uma mensagem de aviso para o usuário.
    else:
        st.warning('Por favor insira as Planilhas em "Importar planilha base".')

def page_preenche_taxa_resposta():

    # Crio uma lista com todos os arquivos que poderão ser usados na página.
    arquivos_previstos = [
        st.session_state.Dados_geral_SCM, st.session_state.Dados_geral_SCC,
        st.session_state.Dados_geral_SC, st.session_state.Dados_geral_SSV
    ]

    arquivos_inseridos = [] # Lista contendo somente os arquivos que de fato foram inseridos pelo usuário.
    sondagens_trabalhadas = [] # Lista contendo as sondagens correspondentes a cada arquivo inserido.
    contador = 0

    # Loop responsável por percorrer todos os arquivos da lista arquivos_previstos e adicionar os que não estão em branco na lista de arquivos_inseridos.
    for arquivo in arquivos_previstos:
        
        if arquivo is not None:
            arquivos_inseridos.append(arquivo)

            # Com base no contador, identifico a sondagem correspondente ao arquivo que foi inserido e a adiciono na lista de sondagens_trabalhadas.
            if contador == 0:
                sondagens_trabalhadas.append('Comércio')
            if contador == 1:
                sondagens_trabalhadas.append('Construção')
            if contador == 2:
                sondagens_trabalhadas.append('Indústria')
            if contador == 3:
                sondagens_trabalhadas.append('Serviços')

        contador += 1

    # Verifico se existem dados para o preenchimento.
    if arquivos_inseridos:

        # Cabeçalho da página.
        st.header('Preencher :red[SACE Sondagem Taxa de Resposta]', divider="red")

        # Espaçamento.
        st.write('')
        st.write('')
        st.write('')
        st.write('')

        container = st.container()
        container.subheader('Importe arquivo a ser atualizado:')

        # Lista onde vou guardar o arquivo a ser modificado pelo usuário.
        uploaded_taxa = container.file_uploader('', accept_multiple_files=False, type=["xlsx"], key="page3")

        # Verifico se a lista tem algum arquivo.
        if uploaded_taxa is not None:

            # Guardo o arquivo passado pelo usuário em formato de df.
            df_uploaded = pd.read_excel(uploaded_taxa)

            # Loop para fazer as modificações de forma recursiva para cada sondagem inserida.
            for i in range(len(sondagens_trabalhadas)):

                # Modifico o valor de df_taxa de acordo com a sondagem que será modificada.
                df_taxa = arquivos_inseridos[i]
                
                # Crio df_atualizado para guardar as informações de df_uploaded modificado pela função preenche_taxa_resposta.
                df_atualizado = dp.preenche_taxa_resposta(df_taxa, df_uploaded, sondagens_trabalhadas[i]) 

                # Redefino o valor de df_uploaded para que na próxima rodagem de cada iteração ele receba as modificações da iteração atual.
                df_uploaded = df_atualizado

            # Salvo o df_atualizado em um buffer de memória.
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df_atualizado.to_excel(writer, index = False, sheet_name="BD_SACE")
            buffer.seek(0)

            # Dou ao usuário uma prévia do conteúdo do arquivo que será baixado por ele.
            st.dataframe(pd.read_excel(buffer), width=800, height=300)

            # Espaçamento.
            st.write('')
            st.write('')
            st.write('')
            st.write('')

            # Permito o download do novo arquivo atualizado.
            st.download_button(
                label="Baixar arquivo atualizado",
                data=buffer,
                file_name= f'BD_SACE_Sondagem_Taxa de Respostas_{dp.dia(dt.date.today())}.{dp.mes(dt.date.today())}.xlsx',
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # Caso não existam dados para o preenchimento, retorno uma mensagem de aviso para o usuário.
    else:
        st.warning('Por favor insira as Planilhas em "Importar planilha base".')

def page_preenche_prioritarias(): 

    # Crio uma lista com todos os arquivos que poderão ser usados na página.   
    arquivos_previstos = [
        st.session_state.Dados_placar_SCM, st.session_state.Dados_placar_SCC,
        st.session_state.Dados_placar_SC, st.session_state.Dados_placar_SSV
    ]

    arquivos_inseridos = [] # Lista contendo somente os arquivos que de fato foram inseridos pelo usuário.
    sondagens_trabalhadas = [] # Lista contendo as sondagens correspondentes a cada arquivo inserido.
    contador = 0

    # Loop responsável por percorrer todos os arquivos da lista arquivos_previstos e adicionar os que não estão em branco na lista de arquivos_inseridos.
    for arquivo in arquivos_previstos:
        
        if arquivo is not None:
            arquivos_inseridos.append(arquivo)

            # Com base no contador, identifico a sondagem correspondente ao arquivo que foi inserido e a adiciono na lista de sondagens_trabalhadas.
            if contador == 0:
                sondagens_trabalhadas.append('Comércio')
            if contador == 1:
                sondagens_trabalhadas.append('Construção')
            if contador == 2:
                sondagens_trabalhadas.append('Indústria')
            if contador == 3:
                sondagens_trabalhadas.append('Serviços')

        contador += 1

    # Verifico se existem dados para o preenchimento.
    if arquivos_inseridos:

        # Cabeçalho da página.
        st.header('Preencher :red[SACE Sondagens Prioritárias]', divider="red")

        # Espaçamento.
        st.write('')
        st.write('')
        st.write('')
        st.write('')

        container = st.container()
        container.subheader('Importe arquivo a ser atualizado:')

        # Lista onde vou guardar o arquivo a ser modificado pelo usuário.
        uploaded_prioritarias = container.file_uploader('', accept_multiple_files=False, type=["xlsx"], key="page4")

        # Verifico se a lista tem algum arquivo.
        if uploaded_prioritarias is not None:

            # Crio uma sessão com uma animação de carregamento enquanto as modificações são geradas.
            with st.spinner('Preenchendo a nova planilha. Por favor aguarde...'):
                
                # Carrego a planilha existente.
                workbook = load_workbook(uploaded_prioritarias)

           
                # Loop para fazer as modificações de forma recursiva para cada sondagem inserida.
                for i in range(len(sondagens_trabalhadas)):

                    # Modifico o valor de df_status de acordo com a sondagem que será modificada.
                    df_placar = arquivos_inseridos[i]

                    # Redefino o valor de workbook, que inicialmente era igual ao arquivo inicial, para o arquivo já modificado pela função preencher_status.
                    workbook = dp.preenche_prioritarias(df_placar, workbook, sondagens_trabalhadas[i]) 

                # Salvo o workbook em um buffer de memória.
                buffer = BytesIO()
                workbook.save(buffer)
                buffer.seek(0)

            # Dou ao usuário uma prévia do conteúdo do arquivo que será baixado por ele.
            st.dataframe(pd.read_excel(buffer), width=800, height=300)

            # Espaçamento.
            st.write('')
            st.write('')
            st.write('')
            st.write('')

            # Permito o download do novo arquivo atualizado.
            st.download_button(
                label="Baixar arquivo atualizado",
                data=buffer,
                file_name= f'BD_SACE_Sondagem_Prioritárias_{dp.dia(dt.date.today())}.{dp.mes(dt.date.today())}.xlsx',
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # Caso não existam dados para o preenchimento, retorno uma mensagem de aviso para o usuário.
    else:
        st.warning('Por favor insira as Planilhas em "Importar planilha base".')