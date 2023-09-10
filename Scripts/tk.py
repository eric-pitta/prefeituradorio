import tkinter as tk
from tkinter import messagebox
from google.oauth2.service_account import Credentials
import gspread
from googleapiclient.discovery import build
from datetime import date
import re

def execute_spreadsheet_operation():
    creds = Credentials.from_service_account_file(r'C:\Users\ericp\sheet-to-doc\envcga\cga-project-392415-5ffea2fff357.json',
                                                    scopes=['https://www.googleapis.com/auth/spreadsheets',
                                                            'https://www.googleapis.com/auth/drive',
                                                            'https://www.googleapis.com/auth/documents'])

    client = gspread.authorize(creds)
   
    spreadsheet_url = entry.get()
    spreadsheet_key = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', spreadsheet_url)
    if spreadsheet_key:
        spreadsheet_key = spreadsheet_key.group(1)
        spreadsheet = client.open_by_key(spreadsheet_key)


    worksheet = spreadsheet.get_worksheet(0)

    print("A planilha foi acessada com sucesso!")

    data = worksheet.get_all_values()

    docs_service = build('docs', 'v1', credentials=creds)
    document = docs_service.documents().create().execute()

    document_id = document['documentId']
    print(f'O documento foi criado com sucesso. O ID do documento é {document_id}.')

    primeira_portaria = int(entry.get())

    data_atual = date.today()
    dia = data_atual.day
    meses_portugues = {1: 'JANEIRO', 2: 'FEVEREIRO', 3: 'MARÇO', 4: 'ABRIL', 5: 'MAIO', 6: 'JUNHO', 7: 'JULHO', 8: 'AGOSTO', 9: 'SETEMBRO', 10: 'OUTUBRO', 11: 'NOVEMBRO', 12: 'DEZEMBRO'}
    mes = meses_portugues[data_atual.month]
    ano = data_atual.year

    ordem = primeira_portaria

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
        bold_content = re.findall(r'(?:Designar|Nomear|Exonerar, a pedido,|Dispensar, a pedido,|Exonerar|Dispensar)(.*?)(?=(,|$))', content)

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
                        'alignment': 'START'
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
    drive_service = build('drive', 'v3', credentials=creds)

    def share_with_me(file_id):
        def callback(request_id, response, exception):
            if exception:
                print(exception)
            else:
                print(response)

        batch = drive_service.new_batch_http_request(callback=callback)
        user_permission = {
            'type': 'user',
            'role': 'writer',
            'emailAddress': 'ericpitta.pcrj@gmail.com'
        }
        batch.add(drive_service.permissions().create(
            fileId=file_id,
            body=user_permission,
            fields='id',
            sendNotificationEmail=False
        ))
        batch.execute()

    share_with_me(document_id)

    messagebox.showinfo("Sucesso", "O script foi executado com sucesso!")


root = tk.Tk()
root.geometry("300x200")

label = tk.Label(root, text="Digite o número da primeira portaria do dia:")
label.pack(pady=10)

entry = tk.Entry(root)
entry.pack(pady=10)


label = tk.Label(root, text="URL da planilha:")
label.pack(pady=10)

entry = tk.Entry(root)
entry.pack(pady=10)




button = tk.Button(root, text="Criar esquema", command=execute_spreadsheet_operation)
button.pack()

root.mainloop()
