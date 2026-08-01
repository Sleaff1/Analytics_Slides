Uma ferramenta leve e modular em C# (.NET) para extrair dados do Google Analytics 4 (GA4) e injetá-los dinamicamente em templates customizados do PowerPoint (.pptx).

O sistema manipula os arquivos via OpenXML, dispensando a instalação do Pacote Office e introduzindo o conceito de "Living Templates", onde o histórico dos gráficos é acumulado de forma automática mês a mês no arquivo base.

[!NOTE]

Este projeto foi desenvolvido para automatizar fechamentos de relatórios mensais e é projetado para rodar de forma agnóstica em qualquer sistema operacional compatível com .NET (Linux, Windows, macOS).

Features
Modelos Vivos (Living Templates): O sistema atualiza diretamente a planilha base (Excel embutido) dos gráficos no arquivo original. Os gráficos acumulam dados sequencialmente sem perder o formato.

Zero Dependências: Manipula a árvore XML dos slides. Não requer licenças ou processos do Microsoft Office rodando em segundo plano.

Orientado a Configuração: Mapeamento de clientes, IDs de propriedade e diretórios gerenciados inteiramente por um arquivo genérico Clientes.json.

Tratamento de I/O: Detecta bloqueios de arquivo (slides abertos pelo usuário final) e pula o processo elegantemente sem interromper a fila de automação.

Dependências
O projeto utiliza os seguintes pacotes NuGet:

Google.Analytics.Data.V1Beta - Cliente oficial da API do GA4.

DocumentFormat.OpenXml - SDK para manipulação estrutural de documentos Office.

System.Text.Json - Para parsing nativo da configuração.

Instalação
1. Clone o repositório:

git clone https://github.com/seu-usuario/ga4-to-pptx.git
cd ga4-to-pptx

Restaure e compile:

Bash
dotnet restore
dotnet build -c Release

Autenticação (Google Cloud)
Crie uma Conta de Serviço (Service Account) no Google Cloud Console.

Ative a Google Analytics Data API.

Gere e faça o download da credencial em formato .json.

Adicione o e-mail da Conta de Serviço recém-criada como Leitor (Viewer) diretamente na aba de administração das propriedades do GA4 desejadas.

Clientes.json
Crie o arquivo de configuração para rotear as entradas e saídas do sistema.

Ao definir os caminhos em ambientes Windows, lembre-se de escapar as barras no JSON (C:\\Users\\...). Em ambientes Unix/Linux, utilize o padrão convencional sem escapes.

JSON
[
  {
    "Nome": "Empresa Alpha",
    "Ga4PropertyId": "123456789",
    "CaminhoTemplateSlide": "/home/samuel/Templates/Template_Alpha.pptx",
    "PastaDestino": "/home/samuel/Relatorios_Mensais"
  }
]

