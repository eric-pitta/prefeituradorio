from google.oauth2.service_account import Credentials
import gspread
from googleapiclient.discovery import build
from datetime import date

creds = Credentials.from_service_account_file(r'C:\Users\ericp\sheet-to-doc\envcga\cga-project-392415-5ffea2fff357.json', 
                                              scopes=['https://www.googleapis.com/auth/spreadsheets',
                                                      'https://www.googleapis.com/auth/drive',
                                                      'https://www.googleapis.com/auth/documents'])

client = gspread.authorize(creds)
spreadsheet = client.open_by_key('1jCarIivkpPuYb9VirTaidJS9EVR9XWmqUkzeUwUTqj0')
worksheet = spreadsheet.get_worksheet(0)

print("A planilha foi acessada com sucesso!")

data = worksheet.get_all_values()

doc_content = ''
docs_service = build('docs', 'v1', credentials=creds)
document = docs_service.documents().create().execute()

document_id = document['documentId']
print(f'O documento foi criado com sucesso. O ID do documento é {document_id}.')

primeira_portaria = int(input("Digite o número da primeira portaria: "))

data_atual = date.today()
dia = data_atual.day
meses_portugues = {1: 'JANEIRO', 2: 'FEVEREIRO', 3: 'MARÇO', 4: 'ABRIL', 5: 'MAIO', 6: 'JUNHO', 7: 'JULHO', 8: 'AGOSTO', 9: 'SETEMBRO', 10: 'OUTUBRO', 11: 'NOVEMBRO', 12: 'DEZEMBRO'}
mes = meses_portugues[data_atual.month]
ano = data_atual.year

ordem = primeira_portaria
doc_length = 0

# Adiciona "SUBSECRETARIA DE GESTÃO" no início do documento
titulo = 'SUBSECRETARIA DE GESTÃO\n\n'
doc_content += titulo


for row in data:
    portaria_text = f'PORTARIA "P" Nº {ordem} DE {dia} DE {mes} DE {ano}\n'
    content = ' '.join(row) + '\n\n'
    doc_content += portaria_text + content

    # Incrementa o número da portaria para a próxima iteração
    ordem += 1

# Insere o texto no documento
request = {
    'insertText': {
        'location': {
            'index': 1
        },
        'text': doc_content
    }
}
docs_service.documents().batchUpdate(documentId=document_id, body={'requests': [request]}).execute()

doc_length += len(doc_content)

# Atualiza o estilo do título da Subsecretaria de Gestão
request = {
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
}
docs_service.documents().batchUpdate(documentId=document_id, body={'requests': [request]}).execute()


# Centraliza o título da Subsecretaria de Gestão
request = {
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
}
docs_service.documents().batchUpdate(documentId=document_id, body={'requests': [request]}).execute()

print(f'O documento foi atualizado com sucesso. O ID do documento é {document_id}.')

# Compartilhe o novo documento com um endereço de e-mail
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
        'emailAddress': 'ericpitta.pcrj@gmail.com'  # Substitua pelo seu e-mail
    }
    batch.add(drive_service.permissions().create(
        fileId=file_id,
        body=user_permission,
        fields='id',
        sendNotificationEmail=False
    ))
    batch.execute()

share_with_me(document_id)  # Compartilhe o documento recém-criado
