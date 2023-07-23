import streamlit as st
from google.oauth2.service_account import Credentials
import gspread
from googleapiclient.discovery import build
from datetime import date
import re
from googleapiclient.errors import HttpError

creds = Credentials.from_service_account_file(r'C:\Users\ericp\sheet-to-doc\envcga\cga-project-392415-5ffea2fff357.json',
                                              scopes=['https://www.googleapis.com/auth/spreadsheets',
                                                      'https://www.googleapis.com/auth/drive',
                                                      'https://www.googleapis.com/auth/documents',
                                                      'https://www.googleapis.com/auth/gmail.send',
                                                      'https://www.googleapis.com/auth/gmail.modify'])

client = gspread.authorize(creds)


st.set_page_config(
    page_title="Gerador - Portarias e Resoluções",
    page_icon="🖋️",
    layout="wide",
    initial_sidebar_state="expanded",
    
)


def page_one():
    st.title(" PORTARIAS ✍️" )
    

    # Solicite o link do usuário
    st.markdown('_OBS: Antes de continuar não esqueça de verificar se a planilha está sendo compartilhada com :a cga-project-serviceaccount@cga-project-392415.iam.gserviceaccount.com_')
    link = st.text_input('Insira o link da planilha:')

    
    # Solicite o email do usuário
    user_email = st.text_input('Seu email: ')

    #Solicita o primeiro número de portaria
    primeira_portaria = st.text_input("Número da primeira Portaria do dia: ")
    

    data = []
    if st.button('Gerar Documento'):
        # Use regex para extrair o ID do link
        ids = re.findall(r'spreadsheets/d/([a-zA-Z0-9-_]+)', link)

        
        # Se um ID foi encontrado, use-o para abrir a planilha
        
        if ids:
            id = ids[0]
            spreadsheet = client.open_by_key(id)
            st.write(f"A planilha com ID {id} foi acessada com sucesso!")
            
            worksheet = spreadsheet.get_worksheet(0)
            
            data = worksheet.get_all_values()
        else:
            st.write("Nenhum ID encontrado no link.")
            st.stop()

        # Código a partir daqui pode acessar a variável "spreadsheet"

        docs_service = build('docs', 'v1', credentials=creds)
        document = docs_service.documents().create().execute()

        document_id = document['documentId']

        st.write(f"Documento criado com sucesso. Link direto: https://docs.google.com/document/d/{document_id}.") 

            

        data_atual = date.today()
        dia = data_atual.day
        meses_portugues = {1: 'JANEIRO', 2: 'FEVEREIRO', 3: 'MARÇO', 4: 'ABRIL', 5: 'MAIO', 6: 'JUNHO', 7: 'JULHO', 8: 'AGOSTO', 9: 'SETEMBRO', 10: 'OUTUBRO', 11: 'NOVEMBRO', 12: 'DEZEMBRO'}
        mes = meses_portugues[data_atual.month]
        ano = data_atual.year

        ordem = primeira_portaria
        ordem = int(ordem)

        titulo = 'SUBSECRETARIA DE GESTÃO\n\n'
        nova_frase = 'A SUBSECRETÁRIA DE GESTÃO, DA SECRETARIA MUNICIPAL DA CASA CIVIL, no uso das atribuições que lhe são conferidas pela legislação em vigor,\n\n'
        resolve = 'RESOLVE\n'

        requests = [{
            'insertText': {
                'location': {
                    'index': 1
                },
                'text': titulo
            }
        },{
            'updateTextStyle': {
                'range': {
                    'startIndex': 1,
                    'endIndex': len(titulo) + 1
                },
                'textStyle': {
                    'bold': True
                },
                'fields': 'bold'
            }
        },{
            'updateParagraphStyle': {
                'range': {
                    'startIndex': 1,
                    'endIndex': len(titulo) + 1
                },
                'paragraphStyle': {
                    'alignment': 'CENTER'
                },
                'fields': 'alignment'
            }
        }]

        document_length = len(titulo)


        for i, row in enumerate(data, start=1):
            portaria_text = f'\nPORTARIA "P" Nº {ordem} DE {dia} DE {mes} DE {ano}\n'
            content = ' '.join(row) + '\n'
            bold_content = re.findall(r'(?:Designar, com validade a partir de \d{1,2}º? de .*? de \d{4},|Exonerar, a pedido, com validade a partir de \d{1,2}º? de .*? de \d{4},|Exonerar, com validade a partir de \d{1,2}º? de .*? de \d{4},|Nomear, com validade a partir de \d{1,2}º? de .*? de \d{4},|Designar|Nomear|Exonerar, a pedido,|Dispensar, a pedido,|Exonerar|Dispensar)(.*?)(?=(,|$))', content)

            requests.append({
                'insertText': {
                    'location': {
                        'index': document_length
                    },
                    'text': portaria_text
                }
            })
            document_length += len(portaria_text)

            requests.append({
                'updateTextStyle': {
                    'range': {
                        'startIndex': document_length - len(portaria_text),
                        'endIndex': document_length
                    },
                    'textStyle': {
                        'bold': True
                    },
                    'fields': 'bold'
                }
            })

            requests.append({
                'updateParagraphStyle': {
                    'range': {
                        'startIndex': document_length - len(portaria_text),
                        'endIndex': document_length
                    },
                    'paragraphStyle': {
                        'alignment': 'CENTER'
                    },
                    'fields': 'alignment'
                }
            })

            requests.append({
                'insertText': {
                    'location': {
                        'index': document_length
                    },
                    'text': nova_frase
                }
            })
            document_length += len(nova_frase)

            bold_part = 'A SUBSECRETÁRIA DE GESTÃO, DA SECRETARIA MUNICIPAL DA CASA CIVIL, '
            non_bold_part = 'no uso das atribuições que lhe são conferidas pela legislação em vigor,\n\n'

            requests.append({
                'updateTextStyle': {
                    'range': {
                        'startIndex': document_length - len(nova_frase),
                        'endIndex': document_length - len(non_bold_part)
                    },
                    'textStyle': {
                        'bold': True
                    },
                    'fields': 'bold'
                }
            })

            requests.append({
                'updateTextStyle': {
                    'range': {
                        'startIndex': document_length - len(non_bold_part),
                        'endIndex': document_length
                    },
                    'textStyle': {
                        'bold': False
                    },
                    'fields': 'bold'
                }
            })

            requests.append({
                'updateParagraphStyle': {
                    'range': {
                        'startIndex': document_length - len(nova_frase),
                        'endIndex': document_length
                    },
                    'paragraphStyle': {
                        'alignment': 'JUSTIFIED'
                    },
                    'fields': 'alignment'
                }
            })

            requests.append({
                'insertText': {
                    'location': {
                        'index': document_length
                    },
                    'text': resolve
                }
            })
            document_length += len(resolve)

            requests.append({
                'updateTextStyle': {
                    'range': {
                        'startIndex': document_length - len(resolve),
                        'endIndex': document_length
                    },
                    'textStyle': {
                        'bold': False
                    },
                    'fields': 'bold'
                }
            })

            requests.append({
                'updateParagraphStyle': {
                    'range': {
                        'startIndex': document_length - len(resolve),
                        'endIndex': document_length
                    },
                    'paragraphStyle': {
                        'alignment': 'START'
                    },
                    'fields': 'alignment'
                }
            })

            requests.append({
                'insertText': {
                    'location': {
                        'index': document_length
                    },
                    'text': content
                }
            })
            document_length += len(content)

            requests.append({
                'updateTextStyle': {
                    'range': {
                        'startIndex': document_length - len(content),
                        'endIndex': document_length
                    },
                    'textStyle': {
                        'bold': False
                    },
                    'fields': 'bold'
                },
        #-------------------------------
            })
            requests.append({
                'updateParagraphStyle': {
                    'range': {
                        'startIndex': document_length - len(content),
                        'endIndex': document_length
                    },
                    'paragraphStyle': {
                        'alignment': 'JUSTIFIED'
                    },
                    'fields': 'alignment'
                }

        #---------------------------------

            })

            for item in bold_content:
                start_index = document_length - len(content) + content.index(item[0])
                end_index = start_index + len(item[0])
                requests.append({
                    'updateTextStyle': {
                        'range': {
                            'startIndex': start_index,
                            'endIndex': end_index
                        },
                        'textStyle': {
                            'bold': True
                        },
                        'fields': 'bold'
                    }
                })

            if i % 100 == 0:
                docs_service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()
                requests = []

            ordem += 1

        if requests:
            docs_service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()


            def share_with_me(document_id):
                try:
                    drive_service = build('drive', 'v3', credentials=creds)
                    drive_service.permissions().create(
                        fileId=document_id,
                        body={'type': 'user', 'role': 'writer', 'emailAddress': user_email},
                        fields='id'
                    ).execute()
                    st.write('Documento compartilhado com sucesso!') 
                except HttpError as error:
                    st.error(f'Ocorreu um erro ao compartilhar o documento: {error}')

            def rename_file(document_id, new_name):
                try:
                    drive_service = build('drive', 'v3', credentials=creds)
                    drive_service.files().update(
                        fileId=document_id,
                        body={"name": new_name}
                    ).execute()
                    st.write(f'Nome do documento: {new_name}') 
                    st.image('https://i.gifer.com/WtX6.gif')
                except HttpError as error:
                    st.error(f"Ocorreu um erro ao renomear o documento: {error}")

        new_file_name = f"PORTARIAS P Nº {primeira_portaria} A {ordem - 1}"
        share_with_me(document_id)
        rename_file(document_id, new_file_name)

        

def page_two():
    st.title(" RESOLUÇÕES ✒️")
    
    # Solicite o link do usuário
    st.markdown('_OBS: Antes de continuar não esqueça de verificar se a planilha está sendo compartilhada com :a cga-project-serviceaccount@cga-project-392415.iam.gserviceaccount.com_')
    link = st.text_input('Insira o link da planilha:')

    
    # Solicite o email do usuário
    user_email = st.text_input('Seu email: ')

    #Solicita o primeiro número de resolução
    primeira_resolucao = st.text_input("Número da primeira Resolucao do dia: ")
    

    data = []

    

    if st.button('Gerar Documento'):
        
        
        # Use regex para extrair o ID do link
        ids = re.findall(r'spreadsheets/d/([a-zA-Z0-9-_]+)', link)

        
        # Se um ID foi encontrado, use-o para abrir a planilha
        
        if ids:
            id = ids[0]
            spreadsheet = client.open_by_key(id)
            st.write(f"A planilha com ID {id} foi acessada com sucesso!")
            
            worksheet = spreadsheet.get_worksheet(0)
            
            data = worksheet.get_all_values()
        else:
            st.write("Nenhum ID encontrado no link.")
            st.stop()

        # Código a partir daqui pode acessar a variável "spreadsheet"

        docs_service = build('docs', 'v1', credentials=creds)
        document = docs_service.documents().create().execute()

        document_id = document['documentId']

        st.write(f"Documento criado com sucesso. Link direto: https://docs.google.com/document/d/{document_id}.") 

        data_atual = date.today()
        dia = data_atual.day
        meses_portugues = {1: 'JANEIRO',
                        2: 'FEVEREIRO',
                        3: 'MARÇO',
                        4: 'ABRIL',
                        5: 'MAIO',
                        6: 'JUNHO',
                        7: 'JULHO',
                        8: 'AGOSTO',
                        9: 'SETEMBRO',
                        10: 'OUTUBRO',
                        11: 'NOVEMBRO',
                        12: 'DEZEMBRO'}
        mes = meses_portugues[data_atual.month]
        ano = data_atual.year

        ordem = primeira_resolucao
        ordem = int(ordem)


        nova_frase = 'O SECRETÁRIO MUNICIPAL DA CASA CIVIL, no uso das atribuições que lhe são conferidas pela legislação em vigor,\n\n'
        resolve = 'RESOLVE\n'


        requests = []
        document_length = 1


        for i, row in enumerate(data, start=1):
            resolucao_text = f'\nRESOLUÇÃO "P" Nº {ordem} DE {dia} DE {mes} DE {ano}\n'
            content = ' '.join(row) + '\n'
            bold_content = re.findall(r'(?:Designar, com validade a partir de \d{1,2}º? de .*? de \d{4},|Exonerar, a pedido, com validade a partir de \d{1,2}º? de .*? de \d{4},|Exonerar, com validade a partir de \d{1,2}º? de .*? de \d{4},|Nomear, com validade a partir de \d{1,2}º? de .*? de \d{4},|Designar|Nomear|Exonerar, a pedido,|Dispensar, a pedido,|Exonerar|Dispensar)(.*?)(?=(,|$))', content)

            # bold_content = re.findall(r'(?:Designar|Nomear|Exonerar, a pedido,|Dispensar, a pedido,|Exonerar|Dispensar)(.*?)(?=(,|$))', content)



            requests.append({
                'insertText': {
                    'location': {
                        'index': document_length
                    },
                    'text': resolucao_text
                }
            })
            document_length += len(resolucao_text)

            requests.append({
                'updateTextStyle': {
                    'range': {
                        'startIndex': document_length - len(resolucao_text),
                        'endIndex': document_length
                    },
                    'textStyle': {
                        'bold': True
                    },
                    'fields': 'bold'
                }
            })

            requests.append({
                'updateParagraphStyle': {
                    'range': {
                        'startIndex': document_length - len(resolucao_text),
                        'endIndex': document_length
                    },
                    'paragraphStyle': {
                        'alignment': 'CENTER'
                    },
                    'fields': 'alignment'
                }
            })

            requests.append({
                'insertText': {
                    'location': {
                        'index': document_length
                    },
                    'text': nova_frase
                }
            })
            document_length += len(nova_frase)

            bold_part = 'A SUBSECRETÁRIA DE GESTÃO, DA SECRETARIA MUNICIPAL DA CASA CIVIL, '
            non_bold_part = 'no uso das atribuições que lhe são conferidas pela legislação em vigor,\n\n'

            requests.append({
                'updateTextStyle': {
                    'range': {
                        'startIndex': document_length - len(nova_frase),
                        'endIndex': document_length - len(non_bold_part)
                    },
                    'textStyle': {
                        'bold': True
                    },
                    'fields': 'bold'
                }
            })

            requests.append({
                'updateTextStyle': {
                    'range': {
                        'startIndex': document_length - len(non_bold_part + '1'),
                        'endIndex': document_length
                    },
                    'textStyle': {
                        'bold': False
                    },
                    'fields': 'bold'
                }
            })

            requests.append({
                'updateParagraphStyle': {
                    'range': {
                        'startIndex': document_length - len(nova_frase),
                        'endIndex': document_length
                    },
                    'paragraphStyle': {
                        'alignment': 'START'
                    },
                    'fields': 'alignment'
                }
            })

            requests.append({
                'insertText': {
                    'location': {
                        'index': document_length
                    },
                    'text': resolve
                }
            })
            document_length += len(resolve)

            requests.append({
                'updateTextStyle': {
                    'range': {
                        'startIndex': document_length - len(resolve),
                        'endIndex': document_length
                    },
                    'textStyle': {
                        'bold': False
                    },
                    'fields': 'bold'
                }
            })

            requests.append({
                'updateParagraphStyle': {
                    'range': {
                        'startIndex': document_length - len(resolve),
                        'endIndex': document_length
                    },
                    'paragraphStyle': {
                        'alignment': 'START'
                    },
                    'fields': 'alignment'
                }
            })

            requests.append({
                'insertText': {
                    'location': {
                        'index': document_length
                    },
                    'text': content
                }
            })
            document_length += len(content)

            requests.append({
                'updateTextStyle': {
                    'range': {
                        'startIndex': document_length - len(content),
                        'endIndex': document_length
                    },
                    'textStyle': {
                        'bold': False
                    },
                    'fields': 'bold'
                },
        #-------------------------------
            })
            requests.append({
                'updateParagraphStyle': {
                    'range': {
                        'startIndex': document_length - len(content),
                        'endIndex': document_length
                    },
                    'paragraphStyle': {
                        'alignment': 'JUSTIFIED'
                    },
                    'fields': 'alignment'
                }

        #---------------------------------

            })

            for item in bold_content:
                start_index = document_length - len(content) + content.index(item[0])
                end_index = start_index + len(item[0])
                requests.append({
                    'updateTextStyle': {
                        'range': {
                            'startIndex': start_index,
                            'endIndex': end_index
                        },
                        'textStyle': {
                            'bold': True
                        },
                        'fields': 'bold'
                    }
                })

            if i % 100 == 0:
                docs_service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()
                requests = []

            ordem += 1

        if requests:
            docs_service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()
            

        

            def share_with_me(document_id):
                try:
                    drive_service = build('drive', 'v3', credentials=creds)
                    drive_service.permissions().create(
                        fileId=document_id,
                        body={'type': 'user', 'role': 'writer', 'emailAddress': user_email},
                        fields='id'
                    ).execute()
                    st.write('Documento compartilhado com sucesso!') 
                except HttpError as error:
                    st.error(f'Ocorreu um erro ao compartilhar o documento: {error}')

            def rename_file(document_id, new_name):
                try:
                    drive_service = build('drive', 'v3', credentials=creds)
                    drive_service.files().update(
                        fileId=document_id,
                        body={"name": new_name}
                    ).execute()
                    st.write(f'Nome do documento: {new_name}') 
                    st.image('https://tenor.com/pt-BR/view/cristiano-ronaldo-ronaldo-manchester-ronaldo-united-ronaldo-manchester-united-cristiano-ronaldo-manchester-united-gif-23658163.gif')
                except HttpError as error:
                    st.error(f"Ocorreu um erro ao renomear o documento: {error}")

        new_file_name = f"RESOLUÇÕES P Nº {primeira_resolucao} A {ordem - 1}"
        share_with_me(document_id)
        rename_file(document_id, new_file_name)

    

PAGES = {
    "Portarias": page_one,
    "Resoluções": page_two
}


def main():
    st.sidebar.title('Crie arquivos para publicação de forma automática')
    choice = st.sidebar.radio("Escolha entre:", list(PAGES.keys()))

    PAGES[choice]()

    st.sidebar.markdown("Desenvolvido pela Coordenadoria Geral de Administração do Gabinete do Prefeito - GP/CGA")

def hide_streamlit_style():
    st.markdown("""
        <style>
        MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        </style>
        """, unsafe_allow_html=True)

hide_streamlit_style()

if __name__ == "__main__":
    main()


