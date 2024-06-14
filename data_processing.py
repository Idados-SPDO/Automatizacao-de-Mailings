import streamlit as st
from openpyxl import load_workbook
import pandas as pd
import datetime as dt
from babel.dates import format_date

# Funções para carregamento dos dados iniciais

def load_mailing_geral(content_file=None):
    if content_file is not None:
        content = pd.read_excel(content_file, sheet_name="Mailing Geral")

        return content
    
def load_mailing_status(content_file=None):
    if content_file is not None:
        content = pd.read_excel(content_file, sheet_name="Status")
        
        return content
    
def load_mailing_placar(content_file=None):
    if content_file is not None:

        # Guardo a data de referência para abrir a planilha de 'placar' no dia e mês da referência inputada
        data_preenchimento = st.session_state.data_referencia

        content = pd.read_excel(content_file, sheet_name=f'Placar {dia(data_preenchimento)}.{mes(data_preenchimento)}', header=None)
        
        return content
    
def load_sondagem_status(content_file=None):
    if content_file is not None:
        content = pd.read_excel(content_file, sheet_name="Sondagem_Status")

# Funções de preenchimento de planilhas

def preenche_indicador(df_status, workbook, sheet):

    # Guardo a data de referência para preencher a coluna "data de preenchimento"
    data_preenchimento = st.session_state.data_referencia
    
    # Carrego a sheet da planilha existente
    sheet_workbook = workbook[sheet]

    # Itero sobre as linhas do DataFrame e atualizo as células na planilha (nome da sheet)
    for indice, linha in df_status.iterrows():

        # Encontro o índice da linha correspondente na planilha
        for row in sheet_workbook.iter_rows(min_row=2): 

            # Verifico se o valor de da linha na coluna "Status" do df está na coluna 2 da planilha   
            if linha['STATUS'].strip() in row[2].value:

                # Atualizo os valores nas colunas desejadas
                sheet_workbook.cell(row=row[0].row, column=1, value= f'{dia(data_preenchimento)}/{mes(data_preenchimento)}/{ano(data_preenchimento)}')
                sheet_workbook.cell(row=row[0].row, column=4, value=linha['QUANTIDADE'])
                sheet_workbook.cell(row=row[0].row, column=5, value=linha['%'])
                break

    return workbook

def preenche_status(df_status, workbook, sondagem): 

    # dicionário de equivalencias de status para grupo de status
    status_to_group = {
        'CONCLUÍDA': 'REALIZADO',
        'AGENDAMENTO': 'EM NEGOCIAÇÃO',
        'E-MAIL ENVIADO': 'EM NEGOCIAÇÃO',
        'WHATSAPP ENVIADO': 'EM NEGOCIAÇÃO',
        'NÃO ENCONTRADO': 'EM NEGOCIAÇÃO',
        'EM NEGOCIAÇÃO': 'EM NEGOCIAÇÃO',
        'AINDA NÃO TRABALHADO': 'PENDENTE',
        'RECUSA NO MÊS': 'SEM SUCESSO',
        'TELEFONE ERRADO': 'SEM SUCESSO',
        'PROBLEMA NO TELEFONE': 'SEM SUCESSO',
        'NÃO DESEJA MAIS PARTICIPAR': 'DESATIVAR',
        'SEM PERFIL': 'DESATIVAR',
        'EMPRESA FECHADA': 'EMP_FECHADA',
        'MAILING EXCEDENTE': 'BASE EXTRA',
        'NÃO RESPONDE HÁ MESES, EM TRATAMENTO': 'RECUPERAÇÃO',
        'EM ESTUDO': 'RECUPERAÇÃO',
        'EMPRESA RECUPERADA/ PROSPECTADA': 'PROSPECTADO RECUPERADO'
    }

    # Adiciono a nova coluna 'GRUPO DE STATUS' usando o mapeamento
    df_status['GRUPO DE STATUS'] = df_status['STATUS'].str.strip().map(status_to_group).fillna('')

    # Reordeno as colunas
    nova_ordem = ['GRUPO DE STATUS', 'STATUS', 'QUANTIDADE', '%']
    df_status = df_status[nova_ordem]

    # Remove as colunas 'STATUS' e '%'
    df_status = df_status.drop(columns=['STATUS', '%'])

     # Removo as linhas onde 'GRUPO DE STATUS' está vazio
    df_status = df_status[df_status['GRUPO DE STATUS'] != '']

    # Agrupo por 'GRUPO DE STATUS' e soma os valores de 'QUANTIDADE'
    df_status = df_status.groupby('GRUPO DE STATUS', as_index=False).sum()

    # Leio o workbook na sheet de interesse   
    sheet_worbook = workbook['BD_SACE_Sondagem_Status']

    # Variável utilizada para preencher a coluna de REFERÊNCIA (método que encontrei para fazer o preenchimento sem dor de cabeça relacionado a conflito de formatos de data)
    data_atual = f'{dia(st.session_state.data_referencia)}/{mes(st.session_state.data_referencia)}/{ano(st.session_state.data_referencia)}'

    # Itero sobre os grupos únicos de STATUS no DataFrame
    for indice, linha in df_status.iterrows():

        # Verifico a qual classe de sondagem a linha pertence e monto a nova linha a ser adicionada
        if sondagem == 'COMÉRCIO':
            new_row = [data_atual, sondagem, 800, linha['GRUPO DE STATUS'], linha['QUANTIDADE']]

        elif sondagem == 'CONSTRUÇÃO':
            new_row = [data_atual, sondagem, 700, linha['GRUPO DE STATUS'], linha['QUANTIDADE']]

        elif sondagem == 'INDÚSTRIA':
            new_row = [data_atual, sondagem, 1150, linha['GRUPO DE STATUS'], linha['QUANTIDADE']]

        elif sondagem == 'SERVIÇOS':
            new_row = [data_atual, sondagem, 1700, linha['GRUPO DE STATUS'], linha['QUANTIDADE']]

        else:
            continue
        
        # Insiro os valores na próxima linha vazia
        sheet_worbook.append(new_row)

    return workbook

def preenche_taxa_resposta(df_taxa, df_uploaded, sondagem):

    # Extraio os valores em branco de 'Placar'
    df_taxa = df_taxa.dropna(subset=['Placar'])

     # Converto a coluna 'Placar' para o tipo de dados de data
    df_taxa['Placar'] = pd.to_datetime(df_taxa['Placar'], format="%d/%m/%Y")

    # Função para determinar o valor da nova coluna 'Ação'
    def determinar_acao(row):

        # Variáveis de data iniciais 
        min_date = df_taxa['Placar'].min()
        row_date = row['Placar']

        # Variáveis de data no formato que preciso para fazer a comparação. OBS: optei pelo uso de f-strings devido a dificuldade que estava tendo para comparar datas datetime antes. 
        data_minima = f'{dia(min_date)}/{mes(min_date)}/{ano(min_date)}'
        data_analise = f'{dia(row_date)}/{mes(row_date)}/{ano(row_date)}'
        
        if pd.isnull(row['Placar']):
            return ''
        elif data_analise == data_minima:
            return 'espontânea'
        else:
            return 'esforço equipe'

    # Crio a coluna 'Ação' com base na função
    df_taxa['Ação'] = df_taxa.apply(determinar_acao, axis=1)

    #Adiciono as linhas nas quais serão aplicadas os novos valores e as novas datas
    novas_linhas = [
    
        {'Sondagem': sondagem, 'Porte': 'P', 'Ação': 'esforço equipe'},
        {'Sondagem': sondagem, 'Porte': 'M', 'Ação': 'esforço equipe'},
        {'Sondagem': sondagem, 'Porte': 'G', 'Ação': 'esforço equipe'},
        {'Sondagem': sondagem, 'Porte': 'P', 'Ação': 'espontânea'},
        {'Sondagem': sondagem, 'Porte': 'M', 'Ação': 'espontânea'},
        {'Sondagem': sondagem, 'Porte': 'G', 'Ação': 'espontânea'}

    ]

    # Concateno o df_uploaded com o df_taxa modificado
    df_uploaded = pd.concat([df_uploaded, pd.DataFrame(novas_linhas)], ignore_index=True)

    # Crio um filtro no qual a coluna 'Sondagem' é o nome da sondagem atual e a coluna 'Data' é vazio.
    filtro_sondagem = (
        (df_uploaded['Sondagem'] == sondagem) & 
        (df_uploaded['Data'].isnull())
    )

    # Crio uma variável para armazenar a data máxima de df_taxa
    data_atual = f'{dia(st.session_state.data_referencia)}/{mes(st.session_state.data_referencia)}/{ano(st.session_state.data_referencia)}'
    data_atual = pd.to_datetime(data_atual, format='%d/%m/%Y')

    # Aplico o filtro criado anteriormente a df_taxa_resposta
    df_taxa_resposta = df_uploaded[filtro_sondagem]

    # Defino a variável 'indices_validos' como uma lista contendo todos os índices de df_taxa_resposta
    indices_taxa_resposta = df_taxa_resposta.index.tolist()

    # Defino a coluna de 'Data' como a maior data de df_taxa_resposta
    for indice in indices_taxa_resposta:
        data_atual = pd.to_datetime(data_atual)
        df_taxa_resposta.loc[indice, 'Data'] = data_atual.strftime('%d/%m/%Y') 

    # Agrupo df_taxa por 'Ação' e 'Porte' e conto o número de ocorrências
    contagem_por_grupo = df_taxa.groupby(['Ação', 'Porte']).size().reset_index(name='Quantidade')

    # Faço um merge dos resultados de volta no df_taxa_resposta com base nas colunas 'Ação' e 'Porte'
    df_taxa_resposta = pd.merge(df_taxa_resposta, contagem_por_grupo, on=['Ação', 'Porte'], how='left', suffixes=('_df_taxa_resposta', '_contagem_por_grupo'))

    # Preencho valores nulos (caso não haja correspondência)
    df_taxa_resposta['Quantidade'] = df_taxa_resposta['Quantidade_contagem_por_grupo'].fillna(0).astype(int)

    # Removo as colunas extras criadas
    df_taxa_resposta = df_taxa_resposta.drop(['Quantidade_df_taxa_resposta', 'Quantidade_contagem_por_grupo'], axis=1)

    # Removo linhas duplicadas de df_taxa_resposta
    df_taxa_resposta = df_taxa_resposta.drop_duplicates(subset=['Ação', 'Porte'])

    # Converto a coluna 'Data' para o tipo de dados de data
    df_uploaded['Data'] = pd.to_datetime(df_uploaded['Data'] ).dt.strftime('%d/%m/%Y')

    # Identifico as linhas a serem atualizadas
    filtro_sondagem = (df_uploaded['Sondagem'] == sondagem) & ((df_uploaded['Data'].isin(df_taxa_resposta['Data'])) | (df_uploaded['Data'].isnull()))
    linhas_para_atualizar = df_uploaded[filtro_sondagem].head(6)

    # Verifico se há linhas para atualizar
    if not linhas_para_atualizar.empty:

        # Atualizo as linhas específicas com os valores de df_taxa_resposta
        colunas_para_atualizar = ['Data', 'Quantidade'] 

        # Itero sobre os índices e atualizo as linhas correspondentes
        for x, (idx, row) in enumerate(linhas_para_atualizar.iterrows()):
            df_uploaded.loc[idx, colunas_para_atualizar] = df_taxa_resposta.loc[x, colunas_para_atualizar].values

        return df_uploaded
    
def preenche_prioritarias(df_placar, workbook, sondagem):
    
    # Carrego a sheet da planilha existente
    sheet_workbook = workbook['Base']

    # Defino a renomeação de df_geral_placar
    placar_renomeacao = {0:'Código', 1:'Nome', 2:'Data' , 3:'Tipo'}

    # Aplico a renomeação na coluna 'STATUS'
    df_placar = df_placar.rename(columns = placar_renomeacao)

    # Encontro o índice da linha correspondente na planilha
    for row in sheet_workbook.iter_rows(min_row=2): 

        # Verifico se o valor da linha está na coluna 'Código' de df_placar  
        if row[2].value in df_placar['Código'].values:

            # Encontro o índice da correspondência em df_placar
            i_corresp = df_placar[df_placar['Código'] == row[2].value].index[0]

            # Atualizo os valores nas colunas desejadas
            sheet_workbook.cell(row = row[0].row, column = 1, value = df_placar['Data'][i_corresp])
            sheet_workbook.cell(row = row[0].row, column = 2, value = sondagem)
            sheet_workbook.cell(row = row[0].row, column = 6, value = 'sim')

    return workbook

# Demais funções que utilizo para inteligência de tempo

def dia(data):

    dia_atual = dt.datetime.strftime(data, '%d')

    return dia_atual

def mes(data):

    mes_atual = dt.datetime.strftime(data, '%m')

    return mes_atual

def nome_mes(data):

    nome = format_date(data, 'MMMM', locale='pt_BR').capitalize()

    return nome

def ano(data):

    ano_atual = dt.datetime.strftime(data, '%Y')

    return ano_atual

