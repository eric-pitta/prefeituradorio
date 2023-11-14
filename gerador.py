import streamlit as st
from google.oauth2.service_account import Credentials
import os
import gspread
from googleapiclient.discovery import build
from datetime import date
import re
from googleapiclient.errors import HttpError

all_env_vars = st.secrets["DEFAULT"]

# Substitui a sequência de escape '\n' com a quebra de linha real
private_key = all_env_vars["PRIVATE_KEY"].replace("\\n", "\n")

# Define os escopos de autorização
scopes =['https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive',
        'https://www.googleapis.com/auth/documents',
        'https://www.googleapis.com/auth/gmail.send',
        'https://www.googleapis.com/auth/gmail.modify']

# Usa as credenciais carregadas para criar credenciais do Google API
creds = Credentials.from_service_account_info({
    "type": all_env_vars["TYPE"],
    "project_id": all_env_vars["PROJECT_ID"],
    "private_key_id": all_env_vars["PRIVATE_KEY_ID"],
    "private_key": private_key,
    "client_email": all_env_vars["CLIENT_EMAIL"],
    "client_id": all_env_vars["CLIENT_ID"],
    "auth_uri": all_env_vars["AUTH_URI"],
    "token_uri": all_env_vars["TOKEN_URI"],
    "auth_provider_x509_cert_url": all_env_vars["AUTH_PROVIDER_X509_CERT_URL"],
    "client_x509_cert_url": all_env_vars["CLIENT_X509_CERT_URL"],
    "universe_domain": all_env_vars["UNIVERSE_DOMAIN"]
}, scopes=scopes)


client = gspread.authorize(creds)


#Configuração geral do layout da página'
st.set_page_config(
    page_title="Gerador - Portarias e Resoluções",
    page_icon="🖋️",
    layout="wide",
    initial_sidebar_state="expanded",
    
)

#Definição conteúdo primeira página
def page_one():
    st.title(" PORTARIAS ✍️" )
    

    # Solicita o link do usuário
    st.markdown('''_OBS: Antes de continuar não esqueça de verificar se a planilha está sendo compartilhada com o email: 102277867499-compute@developer.gserviceaccount.com'
''')

    link = st.text_input('Insira o link da planilha:')

    
    # Solicita o email do usuário
    user_email = st.text_input('Seu email: ')

    #Solicita o primeiro número de portaria
    primeira_portaria = st.text_input("Número da primeira Portaria do dia: ")
    
    
    data = []
    #Criação botão que executa o script
    if st.button('Gerar Documento'):
        # Use regex para extrair o ID do link
        ids = re.findall(r'spreadsheets/d/([a-zA-Z0-9-_]+)', link)

        
        # Se um ID é encontrado ele é usado para abrir a planilha        
        if ids:
            id = ids[0]
            spreadsheet = client.open_by_key(id)
            st.write("Planilha acessada com sucesso! ✅")

            #o 0 se refere a primeira aba
            worksheet = spreadsheet.get_worksheet(0)

            #pega todos os dados da planilha e aba selecionada
            data = worksheet.get_all_values()
        else:
            st.write("Nenhum ID encontrado no link.")
            st.stop()

        # função build da biblioteca google-api-python-client é para criar um objeto de serviço que você pode usar para interagir com a API do Google Docs
        # 'docs' especifica que desejamos acessar a API do Google Docs, 'v1' é a versão da API, e credentials=creds são as credenciais para autenticar a aplicação.
        docs_service = build('docs', 'v1', credentials=creds)

        # objeto de serviço recém criado para fazer uma chamada à API do Google Docs e criar um novo documento
        # a função documents() retorna uma referência à coleção de documentos na API do Google Docs 
        # a função create() inicia uma solicitação para criar um novo documento 
        # execute() envia a solicitação e retorna a resposta.
        document = docs_service.documents().create().execute()

        # Armazena o ID único do novo documento Google criado para futuras solicitações à API
        document_id = document['documentId']

        # exibe a mensagem no app de que o doc foi criado com sucesso
        st.write("Documento criado com sucesso! ✅") 

            
        #arranjo das datas com a lib datetime para posterior inclusão nos atos
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

        #cria variavel ordem que será incrementada sempre +1, posteriormente, para que os atos sigam uma ordem numérica
        ordem = primeira_portaria
        ordem = int(ordem)

        #aqui começa estruturação da formatação do documento
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

        #laço for que vai buscar em cada célula da planilha pelos padrões regex determinados em bold_content, pois tudo o que ele encontrar deve ser colocado em negrito (nome das pessoas)
        for i, row in enumerate(data, start=1):
            portaria_text = f'\nPORTARIA "P" Nº {ordem} DE {dia} DE {mes} DE {ano}\n'
            content = ' '.join(row) + '\n'
            bold_content = re.findall(r'(?:Designar, com validade a partir de \d{1,2}º? de .*? de \d{4},|Exonerar, a pedido, com validade a partir de \d{1,2}º? de .*? de \d{4},|Exonerar, com validade a partir de \d{1,2}º? de .*? de \d{4},|Nomear, com validade a partir de \d{1,2}º? de .*? de \d{4},|Designar|Alterar a alocação de|Nomear|Alocar|Exonerar, a pedido,|Dispensar, a pedido,|Exonerar|Dispensar)(.*?)(?=(,|$))', content)

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

        
            })

            #percorrendo uma lista de substrings que devem ser formatadas em negrito e adicionando as solicitações de atualização correspondentes a uma lista de solicitações,
            #essas solicitações podem então ser enviadas à API do Google Docs para atualizar a formatação do texto no documento.
            for item in bold_content:
                item_upper = (item[0].upper(), item[1])
                start_index = document_length - len(content) + content.index(item[0])
                end_index = start_index + len(item[0])
                 
                # remove o texto original 
                requests.append({
                    'deleteContentRange': {
                        'range': {
                            'startIndex': start_index,
                            'endIndex': end_index
                        }
                    }
                })
                
                # insere o novo texto maiúsculo
                requests.append({
                    'insertText': {
                        'location': {
                            'index': start_index
                        },
                        'text': item_upper[0]
                    }
                })
                
                # atualiza o textstyle
                requests.append({
                    'updateTextStyle': {
                        'range': {
                            'startIndex': start_index,
                            'endIndex': start_index + len(item_upper[0])
                        },
                        'textStyle': {
                            'bold': True
                        },
                        'fields': 'bold'
                    }
                })

            #envio de solicitações de atualização de estilo de texto para a API do Google Docs em lotes de 100, para otimizar a eficiência da comunicação com a API.
            if i % 100 == 0:
                docs_service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()
                requests = []
           
            #incrimenta +1 na sequência dos atos
            ordem += 1
          
        # garante que todas as solicitações de atualização de estilo de texto foram enviadas para a API do Google Docs, 
        # mesmo que o número total de solicitações não seja um múltiplo exato de 100
        # (ou seja, ainda existem solicitações restantes na lista requests após o envio em lotes de 100).
        if requests:
            docs_service.documents().batchUpdate(documentId=document_id, body={'requests': requests}).execute()

            #função responsável por compartilhar o documento com o email (Drive) fornecido pelo usuário
            def share_with_me(document_id):
                try:
                    drive_service = build('drive', 'v3', credentials=creds)
                    drive_service.permissions().create(
                        fileId=document_id,
                        body={'type': 'user', 'role': 'writer', 'emailAddress': user_email},
                        fields='id'
                    ).execute()
                    st.write('Documento compartilhado com sucesso! ✅') 
                except HttpError as error:
                    st.error(f'Ocorreu um erro ao compartilhar o documento: {error}')
                  
            #função responsável por renomear o documento para a formatação exigida
            def rename_file(document_id, new_name):
                try:
                    drive_service = build('drive', 'v3', credentials=creds)
                    drive_service.files().update(
                        fileId=document_id,
                        body={"name": new_name}
                    ).execute()
                    st.write(f'Nome: {new_name}')
                    st.write(f"Link direto: https://docs.google.com/document/d/{document_id}")
                    st.image('https://i.gifer.com/WtX6.gif')
                except HttpError as error:
                    st.error(f"Ocorreu um erro ao renomear o documento: {error}")

        #new_file_name define a forma como o documento será renomeado
        new_file_name = f"PORTARIAS P Nº {primeira_portaria} A {ordem - 1}"
        share_with_me(document_id)
        rename_file(document_id, new_file_name)

        

def page_two():
    st.title(" RESOLUÇÕES ✒️")
    
    # Solicite o link do usuário
    st.markdown('''_OBS: Antes de continuar não esqueça de verificar se a planilha está sendo compartilhada com o email: 102277867499-compute@developer.gserviceaccount.com
''')

    link = st.text_input('Insira o link da planilha:')

    
    # Solicite o email do usuário
    user_email = st.text_input('Seu email: ')

    #Solicita o primeiro número de resolução
    primeira_resolucao = st.text_input("Número da primeira Resolucao do dia: ")
    

    data = []    

    if st.button('Gerar Documento'):      
        ids = re.findall(r'spreadsheets/d/([a-zA-Z0-9-_]+)', link)        
        
        if ids:
            id = ids[0]
            spreadsheet = client.open_by_key(id)
            st.write("Planilha acessada com sucesso! ✅")
            
            worksheet = spreadsheet.get_worksheet(0)
            
            data = worksheet.get_all_values()
        else:
            st.write("Nenhum ID encontrado no link.")
            st.stop()

        # Código a partir daqui pode acessar a variável "spreadsheet"

        docs_service = build('docs', 'v1', credentials=creds)
        document = docs_service.documents().create().execute()

        document_id = document['documentId']

        st.write(f"Documento criado com sucesso! ✅") 

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
            bold_content = re.findall(r'(?:Designar, com validade a partir de \d{1,2}º? de .*? de \d{4},|Exonerar, a pedido, com validade a partir de \d{1,2}º? de .*? de \d{4},|Exonerar, com validade a partir de \d{1,2}º? de .*? de \d{4},|Nomear, com validade a partir de \d{1,2}º? de .*? de \d{4},|Designar|Alterar a alocação de|Nomear|Alocar|Exonerar, a pedido,|Dispensar, a pedido,|Exonerar|Dispensar)(.*?)(?=(,|$))', content)

            
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

        
            })

            for item in bold_content:
                item_upper = (item[0].upper(), item[1])
                start_index = document_length - len(content) + content.index(item[0])
                end_index = start_index + len(item[0])
                 
                # remove o texto original 
                requests.append({
                    'deleteContentRange': {
                        'range': {
                            'startIndex': start_index,
                            'endIndex': end_index
                        }
                    }
                })
                
                # insere o novo texto maiúsculo
                requests.append({
                    'insertText': {
                        'location': {
                            'index': start_index
                        },
                        'text': item_upper[0]
                    }
                })
                
                # atualiza o textstyle
                requests.append({
                    'updateTextStyle': {
                        'range': {
                            'startIndex': start_index,
                            'endIndex': start_index + len(item_upper[0])
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
                    st.write('Documento compartilhado com sucesso! ✅') 
                except HttpError as error:
                    st.error(f'Ocorreu um erro ao compartilhar o documento: {error}')

            def rename_file(document_id, new_name):
                try:
                    drive_service = build('drive', 'v3', credentials=creds)
                    drive_service.files().update(
                        fileId=document_id,
                        body={"name": new_name}
                    ).execute()
                    st.write(f'Nome: {new_name}') 
                    st.write(f"Link direto: https://docs.google.com/document/d/{document_id}")
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


