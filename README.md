# prefeituradorio

Gerador - Portarias e Resoluções

🎯 SOBRE

Esta é uma aplcação que usa o Streamlit como interface gráfica e visa facilitar a criação de documentos de portarias e resoluções que serão publicados no Diário Oficial.
Portarias "P" e Resoluções "P" são atos administrativos internos que lidam com questões de pessoal, geralmente nomeações/exoneções e designações/dispensas. Esses atos precisam ser publicados no Dário Ofical, para cumprir com os princípios da publicidade e transparência. 

✨ OBJETIVOS

Facilitar e eliminar trabalhos repetitivos durante o processamento de arquivos usados nas publicações do Diário Oficial (Gabinete do Prefeito e Casa Civil). Tendo em vista o modo repetitivo que envolve o recebimento desses arquivos e sua edição/formatação para posterior publicação, resolvi criar uma forma na intenção de poupar esforço e tempo. Ele foi projetado para pegar dados de uma planilha do Google e usar esses dados para criar um documento do Google. 

🚀 TECNOLOGIAS

Nesse projeto serão utilizadas as seguintes tecnologias:

-Python

-Streamlit

-Google Sheets

-Google Docs

🚀 DETALHES TÉCNICOS

O aplicativo usa a biblioteca gspread para interagir com as Planilhas Google e a biblioteca googleapiclient para criar e manipular os Documentos Google. Os dados são extraídos da planilha usando o método get_all_values() do objeto de planilha. Os dados então são inseridos no documento Google usando uma série de solicitações de API ao serviço do Google Docs. A formatação do texto (negrito, alinhamento) é aplicada conforme necessário durante esse processo. No final do processo, o documento é compartilhado com o usuário através do serviço Google Drive e renomeado de acordo com o número da primeira e última portaria/resolução inserida.

🏁 COMO INICIAR O PROJETO

# Clone this project

$ git clone https://github.com/{{YOUR_GITHUB_USERNAME}}/prefeituradorio

$ cd prefeituradorio
