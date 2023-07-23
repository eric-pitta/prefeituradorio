# prefeituradorio
Objetivo é facilitar e eliminar trabalhos repetitivos durante o processamento de arquivos usados nas publicações do Diário Oficial (Gabinete do Prefeito e Casa Civil).

Gerador - Portarias e Resoluções

Esta é uma aplcação que usa o Streamlit como interface gráfica e visa facilitar a criação de documentos de portarias e resoluções que serão publicados no Diário Oficial.

Portarias "P" e Resoluções "P" são atos administrativos internos que lidam com questões de pessoal, geralmente nomeações/exoneções e designações/dispensas. Esses atos precisam ser publicados no Dário Ofical, para cumprir com os princípios da publicidade e transparência.

Tendo em vista o modo repetitivo que envolve o recebimento desses arquivos e sua edição/formatação para posterior publicação, resolvi criar uma forma na intenção de poupar esforço e tempo.

Ele foi projetado para pegar dados de uma planilha do Google e usar esses dados para criar um documento do Google.

Como usar
No aplicativo Streamlit, você será solicitado a inserir o link da planilha do Google que contém os dados de suas portarias ou resoluções.

Nota: Certifique-se de que a planilha esteja compartilhada com cga-project-serviceaccount@cga-project-392415.iam.gserviceaccount.com

Digite seu e-mail e o número da primeira portaria do dia.

Clique no botão "Gerar Documento".

O aplicativo então extrai os dados da planilha, cria um novo documento do Google, insere os dados da planilha no documento formatando de acordo com o padrão estabelecido e compartilha o documento com você.

O link direto para o documento também é exibido no aplicativo Streamlit.

Detalhes técnicos
O aplicativo usa a biblioteca gspread para interagir com as Planilhas Google e a biblioteca googleapiclient para criar e manipular os Documentos Google.

Os dados são extraídos da planilha usando o método get_all_values() do objeto de planilha.

Os dados então são inseridos no documento Google usando uma série de solicitações de API ao serviço do Google Docs. A formatação do texto (negrito, alinhamento) é aplicada conforme necessário durante esse processo.

No final do processo, o documento é compartilhado com o usuário através do serviço Google Drive e renomeado de acordo com o número da primeira e última portaria/resolução inserida.
