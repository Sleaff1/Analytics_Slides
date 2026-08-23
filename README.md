
# Slides Google Analytics

Criador de slides em `C#` obtendo informações da `API do Google analytics` e inserido em um modelo pré feito pelo usuário.

O sistema manipula os arquivos via OpenXML, dispensando a instalação do Pacote Office e introduzindo o conceito de "Templates vivos", onde o histórico de gráficos é acumulado de forma automática mês a mês no JSON do cliente

Ou seja, ideal para gerar um relatório mensal do site da sua empresa ou para fins pessoais.
## Features
**Modelos de slide**: O sistema atualiza diretamente a planilha base *(Excel embutido)* dos gráficos no arquivo original. Os gráficos tem seus dados armazenados e inseridos de acordo com o seu mês e ano no slide.

**Zero Dependências**: Manipula a árvore XML dos slides. Não requer licenças ou processos do Microsoft Office rodando em segundo plano.

**Orientado a Configuração**: Mapeamento de clientes, IDs de propriedade e diretórios gerenciados inteiramente por um arquivo genérico Clientes.json.

**Tratamento de I/O**: Detecta bloqueios de arquivo (slides abertos pelo usuário final) e pula o processo sem interromper a fila de automação.
## Dependencias

**O projeto utiliza os seguintes pacotes NuGet:**

`Google.Analytics.Data.V1Beta` - Cliente oficial da API do GA4.

`DocumentFormat.OpenXml` - SDK para manipulação estrutural de documentos Office.

`System.Text.Json` - Para parsing nativo da configuração.
## Instalação

1. **Clone o repositório:**
```Bash
git clone https://github.com/Sleaff1/Analytics_Slides
```
**Ou apenas utilize a URL, dependedo da maneira que esteja realizando a clonagem**

**Restaure e compile:**

```Bash
dotnet restore
dotnet build -c Release
```

## Autenticação Google Cloud

- Crie uma Conta de Serviço no Google Cloud Console.

- Ative a Google Analytics Data API.

- Gere e faça o download da credencial em formato .json.

- Adicione o e-mail da Conta de Serviço recém-criada como Leitor diretamente na aba de administração das propriedades do GA4 desejadas.

## Caminhos

Alguns caminhos dentro do código devem ser alterados para que tudo funcione da maneira correta, todas elas ficam presentes na classe `Main`

* `string caminhoCredencial   = @"sua_pasta";`: Credencial gerada no **Google cloud console**

* `string caminhoClientesJson = @"sua_pasta";`: JSON dos clientes que será explicado no próximo tópico

* `string caminhoTemplate     = @"sua_pasta";`: Template do slide que será utilizado na geração

* `string pastaDestino        = @"sua_pasta";`: Local onde os slides sejam colocados após geração



## Clientes
Crie o arquivo de configuração JSON para rotear as entradas e saídas do sistema.

Ao definir os caminhos em ambientes Windows, lembre-se de escapar as barras no JSON (C:\\Users\\...). Em ambientes Unix/Linux, utilize o padrão convencional sem escapes.

**Exemplo de como o arquivo deve ser:**

```JSON
  {
    "Nome": "Empresa Alpha",
    "Estado": "SP",
    "Ga4PropertyId": "123456789",
    "CaminhoClientePasta": "/home/samuel/Clientes/Empresa Alpha",
  }
```
Cada cliente deve possuir uma pasta onde vai ser inserida a sua logo e o seu JSON com as informações dos gráficos mês a mês, para cada ano. Exemplo:

```JSON
  {
  "Ano": "2026",
  "Meses": {
    "5": {
      "NomeMes": "MAIO",
      "Sessoes": 200,
      "Desktop": 400,
      "Mobile": 120,
      "PageViews": 3500
    },
    "6": {
      "NomeMes": "JUNHO",
      "Sessoes": 2500,
      "Desktop": 750,
      "Mobile": 900,
      "PageViews": 4500
    }
  }
```