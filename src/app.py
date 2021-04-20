import pandas as pd
import datetime as dt
import streamlit as st
import base64


def recebe_arquivo():
    st.subheader('Busque o arquivo das inscrições no Forms')
    st.write('link: https://urless.in/qhaj1')
    arquivo = st.file_uploader('Inclua aqui o arquivo Excel gerado no Forms', type='xlsx')
    return arquivo

def trata_arquivo(arquivo):
    cursos = pd.read_excel(arquivo)
    #alterando o nome das colunas
    cursos.columns = ['ID', 'Início', 'Conclusão','Email1', 'Nome1','Cursos', 
                  'Nome', 'CPF', 'Empresa', 'CNPJ', 
                  'Cidade', 'UF', 'e-mail', 
                  'Telefone', 'Certificado', 'Indicação',
                  'Sugestões', 'Pagamento', 'aviso1', 'aviso2']
    #exclusão das colunas desnecessárias
    cursos.drop(['Nome1', 'Email1','aviso1', 'aviso2'], axis=1, inplace=True)
    #transformando os cursos em uma lista
    cursos['Cursos'] = cursos['Cursos'].str.split(';')
    #dividindo os cursos em linhas
    cursos = cursos.explode('Cursos').reset_index(drop=True)
    cursos.drop(cursos.query('Cursos == "Proteção Financeira - TURMA EXTRA - 07/04/2021 (13:00h - 17:00h)"').index, inplace=True)
    #separando o horário
    cursos[['Cursos', 'Horário']] = cursos['Cursos'].str.split(' \(', expand=True)
    #dividindo a data do curso
    cursos[['Data', 'Curso']] = cursos['Cursos'].str.split(' - ', expand=True)
    #reorganizando as colunas
    cursos = cursos[['ID', 'Início', 'Conclusão', 'Cursos', 'Data', 'Curso', 'Horário','Nome', 'CPF', 'Empresa', 'CNPJ',
       'Cidade', 'UF', 'e-mail', 'Telefone', 'Certificado', 'Indicação',
       'Sugestões', 'Pagamento']]
    #excluindo uma coluna
    cursos.dropna(subset=['Cursos'], inplace=True)
    #excluindo as linhas sem curso
    cursos.dropna(subset=['Curso'], inplace=True)
    return cursos

def gera_arquivo(cursos):   
    #criando em excel
    data = dt.datetime.now()
    arquivo_final = cursos.to_excel(f'cursos_{data.day}{data.month}{data.year}.xls', encoding='UTF-8')
    return arquivo_final

def download_link(object_to_download, download_filename, download_link_text):
    """
    Generates a link to download the given object_to_download.

    object_to_download (str, pd.DataFrame):  The object to be downloaded.
    download_filename (str): filename and extension of file. e.g. mydata.csv, some_txt_output.txt
    download_link_text (str): Text to display for download link.

    Examples:
    download_link(YOUR_DF, 'YOUR_DF.csv', 'Click here to download data!')
    download_link(YOUR_STRING, 'YOUR_STRING.txt', 'Click here to download your text!')

    """
    if isinstance(object_to_download,pd.DataFrame):
        object_to_download = object_to_download.to_csv(index=False, encoding='ISO-8859-1', sep=';')

    # some strings <-> bytes conversions necessary here
    b64 = base64.b64encode(object_to_download.encode()).decode()

    return f'<a href="data:file/txt;base64,{b64}" download="{download_filename}">{download_link_text}</a>'

      


def main():
    st.header('SCRIPT DE TRATAMENTO DO FORMULÁRIO DE CURSOS')
    arquivo = recebe_arquivo()
    if arquivo:
        st.write('Aguarde um momento enquanto tratamos seu arquivo ... isso pode levar uns minutinhos')
        st.image('https://gifimage.net/wp-content/uploads/2017/10/calculations-gif-6.gif')
        cursos = trata_arquivo(arquivo)
        st.dataframe(cursos)
        st.subheader('Salve o arquivo csv gerado')
        nome_do_arquivo = st.text_input('Digite o nome do arquivo seguido de .csv')
        if st.button('Download Dataframe as CSV'):
            tmp_download_link = download_link(cursos, nome_do_arquivo, 'Faça o download do seu arquivo gerado')
            st.markdown(tmp_download_link, unsafe_allow_html=True)
        


if __name__ == '__main__':
    main()