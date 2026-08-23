using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;
using S = DocumentFormat.OpenXml.Spreadsheet;
using GA = Google.Analytics.Data.V1Beta;

namespace AutomacaoAnalyticsRift
{
    public class Cliente
    {
        public string Nome { get; set; } = "";
        public string Estado { get; set; } = "";
        public string Ga4PropertyId { get; set; } = "";
        
        public string CaminhoClientePasta { get; set; } = "";
        
        [JsonIgnore]
        public string CaminhoLogo => Path.Combine(CaminhoClientePasta, "logo.png");
        
        public string CaminhoHistoricoJson(string ano) =>
            Path.Combine(CaminhoClientePasta, $"historico_{ano}.json");
    }
    
    public class DadosMes
    {
        public string NomeMes  { get; set; } = "";   
        public double Sessoes  { get; set; }
        public double Desktop  { get; set; }
        public double Mobile   { get; set; }
        public double PageViews { get; set; }
    }
    
    public class HistoricoAnual
    {
        public string Ano { get; set; } = "";
        
        public Dictionary<string, DadosMes> Meses { get; set; } = new();
    }

    public class PeriodoRelatorio
    {
        public DateTime DataInicio { get; }
        public DateTime DataFim { get; }
        public string NomeMes { get; }
        public int NumeroMes { get; }
        public string Ano { get; }
        public string DiasFormatados { get; }

        public PeriodoRelatorio()
        {
            DateTime dataAtual = DateTime.Now;
            DateTime primeiroDiaMesAtual = new DateTime(dataAtual.Year, dataAtual.Month, 1);
            DataFim = primeiroDiaMesAtual.AddDays(-1);
            DataInicio = new DateTime(DataFim.Year, DataFim.Month, 1);

            NumeroMes = DataInicio.Month;
            Ano = DataInicio.Year.ToString();
            DiasFormatados = $"{DataInicio:dd} a {DataFim:dd}";

            var culturaBR = CultureInfo.GetCultureInfo("pt-BR");
            NomeMes = DataInicio.ToString("MMMM", culturaBR).ToUpper();
        }
    }

    public class DadosAnalytics
    {
        public PeriodoRelatorio Periodo { get; } = new PeriodoRelatorio();

        public string TotalSessoes        { get; set; } = "0";
        public string TotalUsuarios       { get; set; } = "0";
        public string TotalUsuariosAtivos { get; set; } = "0";
        public string TotalPageViews      { get; set; } = "0";

        public double UsuariosDesktop { get; set; } = 0;
        public double UsuariosMobile  { get; set; } = 0;
        public string PctDesktop      { get; set; } = "0%";
        public string PctMobile       { get; set; } = "0%";

        public List<string> ListaNavegadores { get; } = new List<string>();
        public List<string> ListaResolucoes  { get; } = new List<string>();
        public List<string> ListaCidades     { get; } = new List<string>();
        public List<string> ListaPaginas     { get; } = new List<string>();
    }
    public class AnalyticsService
    {
        private readonly GA.BetaAnalyticsDataClient _clienteGA4;

        public AnalyticsService(string caminhoCredencialJson)
        {
            Environment.SetEnvironmentVariable("GOOGLE_APPLICATION_CREDENTIALS", caminhoCredencialJson);
            _clienteGA4 = GA.BetaAnalyticsDataClient.Create();
        }

        public DadosAnalytics ObterDados(string propertyId)
        {
            var dados = new DadosAnalytics();
            string dataInicio     = dados.Periodo.DataInicio.ToString("yyyy-MM-dd");
            string dataFim        = dados.Periodo.DataFim.ToString("yyyy-MM-dd");
            string propriedadeGA4 = $"properties/{propertyId}";

            Console.WriteLine($"Buscando dados de {dataInicio} a {dataFim}...");

            var respostaGeral = ExecutarConsulta(propriedadeGA4, dataInicio, dataFim, null, "sessions", "totalUsers", "activeUsers", "screenPageViews");
            if (respostaGeral.Rows.Count > 0)
            {
                dados.TotalSessoes        = respostaGeral.Rows[0].MetricValues[0].Value;
                dados.TotalUsuarios       = respostaGeral.Rows[0].MetricValues[1].Value;
                dados.TotalUsuariosAtivos = respostaGeral.Rows[0].MetricValues[2].Value;
                dados.TotalPageViews      = respostaGeral.Rows[0].MetricValues[3].Value;
            }

            double.TryParse(dados.TotalSessoes,        out double totalSessoesNumero);
            double.TryParse(dados.TotalUsuariosAtivos,  out double totalUsuariosNumero);
            double.TryParse(dados.TotalPageViews,       out double totalVisualizacoesNumero);

            var respostaDispositivos = ExecutarConsulta(propriedadeGA4, dataInicio, dataFim, "deviceCategory", "activeUsers");
            foreach (var linha in respostaDispositivos.Rows)
            {
                if (linha.DimensionValues.Count == 0 || linha.MetricValues.Count == 0) continue;

                string tipoDispositivo = linha.DimensionValues[0].Value.ToLower();
                double.TryParse(linha.MetricValues[0].Value, out double quantidadeUsuarios);

                if (tipoDispositivo == "desktop") dados.UsuariosDesktop += quantidadeUsuarios;
                else                              dados.UsuariosMobile  += quantidadeUsuarios;
            }

            double totalDispositivos = dados.UsuariosDesktop + dados.UsuariosMobile;
            if (totalDispositivos > 0)
            {
                dados.PctDesktop = $"{(dados.UsuariosDesktop / totalDispositivos) * 100:F2}%";
                dados.PctMobile  = $"{(dados.UsuariosMobile  / totalDispositivos) * 100:F2}%";
            }

            PreencherLista(dados.ListaNavegadores, propriedadeGA4, dataInicio, dataFim, "browser",           "activeUsers",    7,  totalUsuariosNumero);
            PreencherLista(dados.ListaResolucoes,  propriedadeGA4, dataInicio, dataFim, "screenResolution",  "activeUsers",    10, totalUsuariosNumero);
            PreencherLista(dados.ListaCidades,     propriedadeGA4, dataInicio, dataFim, "city",              "activeUsers",    10, totalUsuariosNumero);
            PreencherLista(dados.ListaPaginas,     propriedadeGA4, dataInicio, dataFim, "pageTitle",         "screenPageViews",10, totalVisualizacoesNumero);

            return dados;
        }

        private GA.RunReportResponse ExecutarConsulta(string propriedade, string dataInicio, string dataFim,
                                                      string? dimensao, params string[] metricas)
        {
            var requisicao = new GA.RunReportRequest
            {
                Property   = propriedade,
                DateRanges = { new GA.DateRange { StartDate = dataInicio, EndDate = dataFim } },
                Limit      = 10000
            };

            if (!string.IsNullOrEmpty(dimensao))
                requisicao.Dimensions.Add(new GA.Dimension { Name = dimensao });

            foreach (var metrica in metricas)
                requisicao.Metrics.Add(new GA.Metric { Name = metrica });

            if (!string.IsNullOrEmpty(dimensao))
            {
                requisicao.OrderBys.Add(new GA.OrderBy
                {
                    Metric = new GA.OrderBy.Types.MetricOrderBy { MetricName = metricas[0] },
                    Desc   = true
                });
            }

            return _clienteGA4.RunReport(requisicao);
        }

        private void PreencherLista(List<string> lista, string propriedade, string dataInicio, string dataFim,
                                    string dimensao, string metrica, int limite, double totalReferencia)
        {
            var resultado = ExecutarConsulta(propriedade, dataInicio, dataFim, dimensao, metrica);

            int posicao = 1;
            foreach (var linha in resultado.Rows.Take(limite))
            {
                if (linha.DimensionValues.Count == 0 || linha.MetricValues.Count == 0) continue;

                string nomeItem   = linha.DimensionValues[0].Value;
                string valorTexto = linha.MetricValues[0].Value;
                double.TryParse(valorTexto, out double valorNumerico);

                double percentual = totalReferencia > 0 ? (valorNumerico / totalReferencia) * 100 : 0;

                lista.Add($"{posicao} {nomeItem} - {valorTexto} ({percentual:F2}%)");
                posicao++;
            }
        }
    }
    public class HistoricoService
    {
        private static readonly JsonSerializerOptions JsonOpts = new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder       = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };
        public HistoricoAnual CarregarOuCriar(string caminhoJson, string ano)
        {
            if (File.Exists(caminhoJson))
            {
                try
                {
                    var textoJson          = File.ReadAllText(caminhoJson);
                    var historicoCarregado = JsonSerializer.Deserialize<HistoricoAnual>(textoJson, JsonOpts);
                    if (historicoCarregado != null) return historicoCarregado;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[AVISO] Erro ao ler historico JSON: {ex.Message}. Criando novo.");
                }
            }

            return new HistoricoAnual { Ano = ano, Meses = new Dictionary<string, DadosMes>() };
        }
        public void SalvarMesAtual(HistoricoAnual historico, DadosAnalytics dados, string caminhoJson)
        {
            string chave = dados.Periodo.NumeroMes.ToString();

            double.TryParse(dados.TotalSessoes,  out double sessoes);
            double.TryParse(dados.TotalPageViews, out double pageViews);

            historico.Meses[chave] = new DadosMes
            {
                NomeMes   = dados.Periodo.NomeMes,
                Sessoes   = sessoes,
                Desktop   = dados.UsuariosDesktop,
                Mobile    = dados.UsuariosMobile,
                PageViews = pageViews
            };

            string? pastaDoArquivo = Path.GetDirectoryName(caminhoJson);
            if (!string.IsNullOrEmpty(pastaDoArquivo) && !Directory.Exists(pastaDoArquivo))
                Directory.CreateDirectory(pastaDoArquivo);

            File.WriteAllText(caminhoJson, JsonSerializer.Serialize(historico, JsonOpts));
            Console.WriteLine($"Historico atualizado: {caminhoJson}");
        }
    }
    public class SlideService
    {
        private readonly HistoricoService _historicoService = new HistoricoService();
        
        
        public void Gerar(Cliente cliente, DadosAnalytics dados, string caminhoTemplate, string pastaDestino)
        {
            if (!Directory.Exists(cliente.CaminhoClientePasta))
                Directory.CreateDirectory(cliente.CaminhoClientePasta);
            
            string caminhoHistorico = cliente.CaminhoHistoricoJson(dados.Periodo.Ano);
            var historicoCliente    = _historicoService.CarregarOuCriar(caminhoHistorico, dados.Periodo.Ano);
            _historicoService.SalvarMesAtual(historicoCliente, dados, caminhoHistorico);
            
            if (!Directory.Exists(pastaDestino))
                Directory.CreateDirectory(pastaDestino);

            string nomeArquivo = $"Relatorio Mensal de acessos ao site - {dados.Periodo.NomeMes} de {dados.Periodo.Ano}" +
                                 $" - {cliente.Nome.Replace(" ", "_")} - {cliente.Estado.Replace(" ", "_")}.pptx";
            string caminhoFinal = Path.Combine(pastaDestino, nomeArquivo);

            Console.WriteLine($"\nIniciando geracao para {cliente.Nome}...");
            
            if (File.Exists(caminhoFinal))
            {
                try
                {
                    File.SetAttributes(caminhoFinal, FileAttributes.Normal);
                    File.Delete(caminhoFinal);
                }
                catch (IOException)
                {
                    Console.WriteLine($"[AVISO] O arquivo {nomeArquivo} esta ABERTO no PowerPoint!");
                    Console.WriteLine("Feche-o para atualizar. Pulando cliente...");
                    return;
                }
            }
            
            File.Copy(caminhoTemplate, caminhoFinal, true);
            
            Console.WriteLine("Atualizando graficos com historico completo...");
            using (PresentationDocument apresentacao = PresentationDocument.Open(caminhoFinal, true))
            {
                var parteApresentacao = apresentacao.PresentationPart!;
                var idsSlides         = parteApresentacao.Presentation.SlideIdList!.Elements<P.SlideId>().ToList();
                
                if (idsSlides.Count > 1)
                {
                    var valoresMensais = ExtrairValores(historicoCliente, m => new[] { m.Sessoes });
                    AtualizarGrafico(parteApresentacao, idsSlides[1], valoresMensais);
                }
                
                if (idsSlides.Count > 2)
                {
                    var valoresMensais = ExtrairValores(historicoCliente, m => new[] { m.Desktop, m.Mobile });
                    AtualizarGrafico(parteApresentacao, idsSlides[2], valoresMensais);
                }
                
                if (idsSlides.Count > 6)
                {
                    var valoresMensais = ExtrairValores(historicoCliente, m => new[] { m.PageViews });
                    AtualizarGrafico(parteApresentacao, idsSlides[6], valoresMensais);
                }

                Console.WriteLine("Injetando textos e listas...");
                foreach (var idSlide in idsSlides)
                {
                    var parteSlide = (SlidePart)parteApresentacao.GetPartById(idSlide.RelationshipId!.Value!);
                    SubstituirTextos(parteSlide, dados);
                    SubstituirListas(parteSlide, dados);
                }

                Console.WriteLine("Substituindo logo...");
                foreach (var idSlide in idsSlides)
                {
                    var parteSlide = (SlidePart)parteApresentacao.GetPartById(idSlide.RelationshipId!.Value!);
                    SubstituirLogo(parteSlide, cliente.CaminhoLogo);
                }
            }

            Console.WriteLine($"Slide concluido e salvo em: {caminhoFinal}");
        }
        
        private List<(int Mes, double[] Valores)> ExtrairValores(
            HistoricoAnual historico, Func<DadosMes, double[]> seletor)
        {
            return historico.Meses
                .Select(entradaMes => (Mes: int.Parse(entradaMes.Key), Valores: seletor(entradaMes.Value)))
                .OrderBy(item => item.Mes)
                .ToList();
        }

        private void SubstituirTextos(SlidePart parteSlide, DadosAnalytics dados)
        {
            var dicionarioSubstituicoes = new Dictionary<string, string>
            {
                { "#MES_NOME#",      dados.Periodo.NomeMes          },
                { "#ANO#",           dados.Periodo.Ano              },
                { "#DIAS_MES#",      dados.Periodo.DiasFormatados   },
                { "#SESSOES#",       dados.TotalSessoes             },
                { "#USUARIOS#",      dados.TotalUsuarios            },
                { "#MOB_PCT#",       dados.PctMobile                },
                { "#DESK_PCT#",      dados.PctDesktop               },
                { "#TOTAL_NAVEG#",   dados.TotalUsuariosAtivos      },
                { "#TOTAL_RESOL#",   dados.TotalUsuariosAtivos      },
                { "#TOTAL_CIDADES#", dados.TotalUsuariosAtivos      },
                { "#TOTAL_PAGINAS#", dados.TotalPageViews           }
            };

            foreach (var textoElemento in parteSlide.Slide.Descendants<A.Text>().Where(x => !string.IsNullOrEmpty(x.Text)))
            {
                foreach (var substituicao in dicionarioSubstituicoes)
                {
                    if (textoElemento.Text.Contains(substituicao.Key))
                        textoElemento.Text = textoElemento.Text.Replace(substituicao.Key, substituicao.Value);
                }
            }
            parteSlide.Slide.Save();
        }

        private void SubstituirListas(SlidePart parteSlide, DadosAnalytics dados)
        {
            ProcessarListaTag(parteSlide, "#NAVEGADORES#", dados.ListaNavegadores);
            ProcessarListaTag(parteSlide, "#RESOLUCOES#",  dados.ListaResolucoes);
            ProcessarListaTag(parteSlide, "#CIDADES#",     dados.ListaCidades);
            ProcessarListaTag(parteSlide, "#PAGINAS#",     dados.ListaPaginas);
            parteSlide.Slide.Save();
        }

        private void ProcessarListaTag(SlidePart parteSlide, string tag, List<string> linhas)
        {
            var textoMarcador = parteSlide.Slide.Descendants<A.Text>()
                .FirstOrDefault(t => t.Text.Contains(tag));
            if (textoMarcador == null) return;

            var execucaoOriginal = textoMarcador.Parent as A.Run;
            var paragrafo        = execucaoOriginal?.Parent as A.Paragraph;
            if (execucaoOriginal == null || paragrafo == null) return;

            if (linhas.Count == 0)
            {
                textoMarcador.Text = "Sem dados no periodo.";
                return;
            }

            OpenXmlElement ultimoElemento = execucaoOriginal;
            for (int i = 0; i < linhas.Count; i++)
            {
                var novaExecucao = (A.Run)execucaoOriginal.CloneNode(true);
                novaExecucao.GetFirstChild<A.Text>()!.Text = linhas[i];
                paragrafo.InsertAfter(novaExecucao, ultimoElemento);
                ultimoElemento = novaExecucao;

                if (i < linhas.Count - 1)
                {
                    var quebradeLinha = new A.Break();
                    paragrafo.InsertAfter(quebradeLinha, ultimoElemento);
                    ultimoElemento = quebradeLinha;
                }
            }
            execucaoOriginal.Remove();
        }

        private void SubstituirLogo(SlidePart parteSlide, string caminhoLogo)
        {
            if (!File.Exists(caminhoLogo))
            {
                Console.WriteLine($"[AVISO] Logo nao encontrada: {caminhoLogo}. Pulando substituicao.");
                return;
            }
            
            P.Picture? imagemPlaceholder = null;
            foreach (var figuraCandidada in parteSlide.Slide.Descendants<P.Picture>())
            {
                var propriedadesVisuais = figuraCandidada.NonVisualPictureProperties?.NonVisualDrawingProperties;
                if (propriedadesVisuais?.Name?.Value == "#LOGO#")
                {
                    imagemPlaceholder = figuraCandidada;
                    break;
                }
            }

            if (imagemPlaceholder == null) return;
            
            string extensaoArquivo = Path.GetExtension(caminhoLogo).ToLowerInvariant();
            string tipoConteudo = extensaoArquivo switch
            {
                ".png"  => "image/png",
                ".jpg"  => "image/jpeg",
                ".jpeg" => "image/jpeg",
                ".gif"  => "image/gif",
                ".bmp"  => "image/bmp",
                ".webp" => "image/webp",
                _       => "image/png"
            };

            ImagePart parteImagem = parteSlide.AddImagePart(tipoConteudo);
            using (var fluxoLogo = File.OpenRead(caminhoLogo))
                parteImagem.FeedData(fluxoLogo);

            string novoIdRelacionamento = parteSlide.GetIdOfPart(parteImagem);

            var elementoImagem = imagemPlaceholder.BlipFill?.Blip;
            if (elementoImagem != null)
                elementoImagem.Embed = novoIdRelacionamento;

            var dimensoes = ObterDimensoesImagem(caminhoLogo);
            if (dimensoes.Width > 0 && dimensoes.Height > 0)
            {
                var transformacao = imagemPlaceholder.ShapeProperties?.Transform2D;
                if (transformacao?.Extents != null && transformacao.Extents.Cx != null && transformacao.Extents.Cy != null)
                {
                    long larguraOriginalEMU = transformacao.Extents.Cx.Value;
                    long alturaOriginalEMU  = transformacao.Extents.Cy.Value;

                    double proporcaoPlaceholder = (double)larguraOriginalEMU / alturaOriginalEMU;
                    double proporcaoImagem      = (double)dimensoes.Width / dimensoes.Height;

                    long novaLargura, novaAltura;

                    if (proporcaoImagem > proporcaoPlaceholder)
                    {
                        novaLargura = larguraOriginalEMU;
                        novaAltura  = (long)(larguraOriginalEMU / proporcaoImagem);
                    }
                    else
                    {
                        novaAltura  = alturaOriginalEMU;
                        novaLargura = (long)(alturaOriginalEMU * proporcaoImagem);
                    }

                    long deslocamentoX = (larguraOriginalEMU - novaLargura) / 2;
                    long deslocamentoY = (alturaOriginalEMU  - novaAltura)  / 2;

                    transformacao.Extents.Cx = novaLargura;
                    transformacao.Extents.Cy = novaAltura;

                    if (transformacao.Offset != null && transformacao.Offset.X != null && transformacao.Offset.Y != null)
                    {
                        transformacao.Offset.X = transformacao.Offset.X.Value + deslocamentoX;
                        transformacao.Offset.Y = transformacao.Offset.Y.Value + deslocamentoY;
                    }
                }
            }

            parteSlide.Slide.Save();
        }

        private (int Width, int Height) ObterDimensoesImagem(string caminhoArquivo)
        {
            try
            {
                using var fluxoArquivo  = new FileStream(caminhoArquivo, FileMode.Open, FileAccess.Read);
                using var leitorBinario = new BinaryReader(fluxoArquivo);

                var assinaturaArquivo = leitorBinario.ReadUInt64();
                if (assinaturaArquivo == 0x0A1A0A0D474E5089)
                {
                    fluxoArquivo.Position = 16;
                    var bytesLargura = leitorBinario.ReadBytes(4); Array.Reverse(bytesLargura);
                    var bytesAltura  = leitorBinario.ReadBytes(4); Array.Reverse(bytesAltura);
                    return (BitConverter.ToInt32(bytesLargura, 0), BitConverter.ToInt32(bytesAltura, 0));
                }

                fluxoArquivo.Position = 0;
                if (leitorBinario.ReadUInt16() == 0xD8FF)
                {
                    while (fluxoArquivo.Position < fluxoArquivo.Length)
                    {
                        byte marcador     = leitorBinario.ReadByte();
                        if (marcador != 0xFF) break;
                        byte tipoMarcador = leitorBinario.ReadByte();

                        if (tipoMarcador >= 0xC0 && tipoMarcador <= 0xC3)
                        {
                            fluxoArquivo.Seek(3, SeekOrigin.Current);
                            var bytesAltura  = leitorBinario.ReadBytes(2); Array.Reverse(bytesAltura);
                            var bytesLargura = leitorBinario.ReadBytes(2); Array.Reverse(bytesLargura);
                            return (BitConverter.ToUInt16(bytesLargura, 0), BitConverter.ToUInt16(bytesAltura, 0));
                        }
                        
                        var bytesTamanho = leitorBinario.ReadBytes(2); Array.Reverse(bytesTamanho);
                        int tamanhoBloco = BitConverter.ToUInt16(bytesTamanho, 0);
                        fluxoArquivo.Seek(tamanhoBloco - 2, SeekOrigin.Current);
                    }
                }
            }
            catch { }
            return (0, 0);
        }

        private void AtualizarGrafico(PresentationPart parteApresentacao, P.SlideId idSlide,
                                      List<(int Mes, double[] Valores)> dadosPorMes)
        {
            SlidePart  parteSlide   = (SlidePart)parteApresentacao.GetPartById(idSlide.RelationshipId!.Value!);
            ChartPart? parteGrafico = parteSlide.ChartParts.FirstOrDefault();
            if (parteGrafico == null) return;

            var parteExcel = parteGrafico.EmbeddedPackagePart;
            if (parteExcel != null)
            {
                using Stream               fluxoDados     = parteExcel.GetStream(FileMode.Open, FileAccess.ReadWrite);
                using SpreadsheetDocument  documentoExcel = SpreadsheetDocument.Open(fluxoDados, true);

                var partePlanilha = documentoExcel.WorkbookPart!;
                var planilha      = partePlanilha.Workbook.Descendants<S.Sheet>().FirstOrDefault();
                if (planilha != null)
                {
                    var parteAba      = (WorksheetPart)partePlanilha.GetPartById(planilha.Id!.Value!);
                    var dadosPlanilha = parteAba.Worksheet.Elements<S.SheetData>().FirstOrDefault();

                    if (dadosPlanilha != null)
                    {
                        foreach (var (mes, valores) in dadosPorMes)
                        {
                            uint    indiceLinha = (uint)(mes + 1);
                            S.Row? linhaAtual   = dadosPlanilha.Elements<S.Row>()
                                .FirstOrDefault(r => r.RowIndex?.Value == indiceLinha);
                            if (linhaAtual == null)
                            {
                                linhaAtual = new S.Row { RowIndex = new UInt32Value(indiceLinha) };
                                dadosPlanilha.Append(linhaAtual);
                            }

                            string[] colunas = { "B", "C", "D", "E" };
                            for (int i = 0; i < valores.Length && i < colunas.Length; i++)
                            {
                                string  referenciaCelula = $"{colunas[i]}{indiceLinha}";
                                S.Cell? celula           = linhaAtual.Elements<S.Cell>()
                                    .FirstOrDefault(c => c.CellReference?.Value == referenciaCelula);
                                if (celula == null)
                                {
                                    celula = new S.Cell { CellReference = new StringValue(referenciaCelula) };
                                    linhaAtual.Append(celula);
                                }
                                celula.CellValue = new S.CellValue(valores[i].ToString(CultureInfo.InvariantCulture));
                                celula.DataType  = new EnumValue<S.CellValues>(S.CellValues.Number);
                            }
                        }
                    }
                }
            }

            var espacoGrafico = parteGrafico.ChartSpace;
            var listaSeries   = espacoGrafico.Descendants<C.SeriesText>().Select(s => s.Parent).ToList();

            foreach (var (mes, valores) in dadosPorMes)
            {
                uint indicePonto = (uint)(mes - 1);

                for (int i = 0; i < listaSeries.Count && i < valores.Length; i++)
                {
                    var referenciaValor = listaSeries[i]!.Descendants<C.NumberReference>().FirstOrDefault();
                    if (referenciaValor?.NumberingCache == null) continue;

                    var pontoAtual = referenciaValor.NumberingCache.Descendants<C.NumericPoint>()
                        .FirstOrDefault(p => p.Index?.Value == indicePonto);

                    if (pontoAtual?.NumericValue != null)
                    {
                        pontoAtual.NumericValue.Text = valores[i].ToString(CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        var novoPonto = new C.NumericPoint { Index = new UInt32Value(indicePonto) };
                        novoPonto.Append(new C.NumericValue(valores[i].ToString(CultureInfo.InvariantCulture)));
                        referenciaValor.NumberingCache.Append(novoPonto);

                        var contadorPontos = referenciaValor.NumberingCache.Descendants<C.PointCount>().FirstOrDefault();
                        if (contadorPontos?.Val != null && contadorPontos.Val.Value <= indicePonto)
                            contadorPontos.Val.Value = indicePonto + 1;
                    }

                    var rotuloAtual = listaSeries[i]!.Descendants<C.DataLabel>()
                        .FirstOrDefault(l => l.Index?.Val?.Value == indicePonto);
                    rotuloAtual?.Descendants<C.ChartText>().FirstOrDefault()?.Remove();
                }
            }

            double maxValorDados = dadosPorMes
                .SelectMany(x => x.Valores)
                .DefaultIfEmpty(0)
                .Max();

            var (maximoEixo, unidadePrincipal) = CalcularEscalaEixo(maxValorDados);

            var eixoVertical = espacoGrafico.Descendants<C.ValueAxis>().FirstOrDefault();
            if (eixoVertical != null)
            {
                var escalonamento = eixoVertical.Descendants<C.Scaling>().FirstOrDefault();
                if (escalonamento != null)
                {
                    var elementoMaximo = escalonamento.Elements<C.MaxAxisValue>().FirstOrDefault();
                    if (elementoMaximo == null)
                    {
                        elementoMaximo = new C.MaxAxisValue();
                        escalonamento.Append(elementoMaximo);
                    }
                    elementoMaximo.Val = maximoEixo;
                }

                var elementoUnidadePrincipal = eixoVertical.Elements<C.MajorUnit>().FirstOrDefault();
                if (elementoUnidadePrincipal == null)
                {
                    elementoUnidadePrincipal = new C.MajorUnit();
                    eixoVertical.Append(elementoUnidadePrincipal);
                }
                elementoUnidadePrincipal.Val = unidadePrincipal;

                var elementoUnidadeSecundaria = eixoVertical.Elements<C.MinorUnit>().FirstOrDefault();
                if (elementoUnidadeSecundaria == null)
                {
                    elementoUnidadeSecundaria = new C.MinorUnit();
                    eixoVertical.Append(elementoUnidadeSecundaria);
                }
                elementoUnidadeSecundaria.Val = unidadePrincipal / 5.0;
            }

            parteGrafico.ChartSpace.Save();
        }
        
        private (double AxisMax, double MajorUnit) CalcularEscalaEixo(double maxValorDados)
        {
            if (maxValorDados <= 0) maxValorDados = 9;

            double maximoBruto = maxValorDados * 1.25;
            
            double ordemGrandeza       = Math.Floor(Math.Log10(maximoBruto));
            double passoArredondamento = 5 * Math.Pow(10, ordemGrandeza - 1);
            
            double maximoEixo = Math.Round(maximoBruto / passoArredondamento) * passoArredondamento;
            if (maximoEixo <= maxValorDados) maximoEixo += passoArredondamento;
            
            double unidadePrincipal = (maximoEixo / passoArredondamento < 4) ? passoArredondamento / 5.0 : passoArredondamento;

            return (maximoEixo, unidadePrincipal);
        }
    }

    class Program
    {
        static void Main()
        {
            string caminhoCredencial   = @"C:\Users\samu0\Documents\Analytics\analytics-automacao-1c71b3e01156.json";
            string caminhoClientesJson = @"C:\Users\samu0\Documents\Analytics\ClientesdaRIFT.json";
            string caminhoTemplate     = @"C:\Users\samu0\Documents\Analytics\Modelo Padrão RIFT.pptx";
            string pastaDestino        = @"C:\Users\samu0\Documents\Analytics\islaide gerado";
            
            if (!File.Exists(caminhoClientesJson))
            {
                Console.WriteLine($"[ERRO] Arquivo de clientes nao encontrado: {caminhoClientesJson}");
                return;
            }

            if (!File.Exists(caminhoTemplate))
            {
                Console.WriteLine($"[ERRO] Template nao encontrado: {caminhoTemplate}");
                return;
            }

            var textoJson = File.ReadAllText(caminhoClientesJson);
            var clientes  = JsonSerializer.Deserialize<List<Cliente>>(textoJson);

            if (clientes == null || clientes.Count == 0)
            {
                Console.WriteLine("[ERRO] A lista de clientes esta vazia ou o JSON esta invalido.");
                return;
            }
            
            var servicoAnalytics = new AnalyticsService(caminhoCredencial);
            var servicoSlide     = new SlideService();

            Console.WriteLine("--- Iniciando Automacao de Relatorios ---");

            foreach (var cliente in clientes)
            {
                try
                {
                    Console.WriteLine($"\nProcessando: {cliente.Nome}");
                    var dados = servicoAnalytics.ObterDados(cliente.Ga4PropertyId);
                    servicoSlide.Gerar(cliente, dados, caminhoTemplate, pastaDestino);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[ERRO] Falha no cliente {cliente.Nome}: {ex.Message}");
                }
            }

            Console.WriteLine("\n--- Processo Finalizado com Sucesso! ---");
        }
    }
}
