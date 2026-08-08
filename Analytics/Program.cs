using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
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
        public string CaminhoTemplateSlide { get; set; } = "";
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
        
        public string TotalSessoes { get; set; } = "0";
        public string TotalUsuarios { get; set; } = "0";     
        public string TotalUsuariosAtivos { get; set; } = "0"; 
        public string TotalPageViews { get; set; } = "0";

        public double UsuariosDesktop { get; set; } = 0;
        public double UsuariosMobile { get; set; } = 0;
        public string PctDesktop { get; set; } = "0%";
        public string PctMobile { get; set; } = "0%";
        
        public List<string> ListaNavegadores { get; } = new List<string>();
        public List<string> ListaResolucoes { get; } = new List<string>();
        public List<string> ListaCidades { get; } = new List<string>();
        public List<string> ListaPaginas { get; } = new List<string>();
    }
    
    public class AnalyticsService
    {
        private readonly GA.BetaAnalyticsDataClient _client;

        public AnalyticsService(string caminhoCredencialJson)
        {
            Environment.SetEnvironmentVariable("GOOGLE_APPLICATION_CREDENTIALS", caminhoCredencialJson);
            _client = GA.BetaAnalyticsDataClient.Create();
        }

        public DadosAnalytics ObterDados(string propertyId)
        {
            var dados = new DadosAnalytics();
            string dtInicio = dados.Periodo.DataInicio.ToString("yyyy-MM-dd");
            string dtFim = dados.Periodo.DataFim.ToString("yyyy-MM-dd");
            string prop = $"properties/{propertyId}";

            Console.WriteLine($"Buscando dados de {dtInicio} a {dtFim}...");

        
            var resGeral = ExecutarConsulta(prop, dtInicio, dtFim, null, "sessions", "totalUsers", "activeUsers", "screenPageViews");
            if (resGeral.Rows.Count > 0)
            {
                dados.TotalSessoes        = resGeral.Rows[0].MetricValues[0].Value; 
                dados.TotalUsuarios       = resGeral.Rows[0].MetricValues[1].Value; 
                dados.TotalUsuariosAtivos = resGeral.Rows[0].MetricValues[2].Value; 
                dados.TotalPageViews      = resGeral.Rows[0].MetricValues[3].Value;
            }
            double.TryParse(dados.TotalSessoes, out double totalSessoesNum);
            double.TryParse(dados.TotalUsuariosAtivos, out double totalUsuariosNum); 
            double.TryParse(dados.TotalPageViews, out double totalViewsNum);
            
            var resDisp = ExecutarConsulta(prop, dtInicio, dtFim, "deviceCategory", "activeUsers");
            foreach (var row in resDisp.Rows)
            {
                if (row.DimensionValues.Count == 0 || row.MetricValues.Count == 0) continue;

                string disp = row.DimensionValues[0].Value.ToLower();
                double.TryParse(row.MetricValues[0].Value, out double val);
                
                if (disp == "desktop") dados.UsuariosDesktop += val;
                else dados.UsuariosMobile += val;
            }

            double totalDisp = dados.UsuariosDesktop + dados.UsuariosMobile;
            if (totalDisp > 0)
            {
                dados.PctDesktop = $"{(dados.UsuariosDesktop / totalDisp) * 100:F2}%";
                dados.PctMobile = $"{(dados.UsuariosMobile / totalDisp) * 100:F2}%";
            }
            
            PreencherLista(dados.ListaNavegadores, prop, dtInicio, dtFim, "browser", "activeUsers", 7, totalUsuariosNum);
            PreencherLista(dados.ListaResolucoes, prop, dtInicio, dtFim, "screenResolution", "activeUsers", 10, totalUsuariosNum);
            PreencherLista(dados.ListaCidades, prop, dtInicio, dtFim, "city", "activeUsers", 10, totalUsuariosNum);
            PreencherLista(dados.ListaPaginas, prop, dtInicio, dtFim, "pageTitle", "screenPageViews", 10, totalViewsNum);

            return dados;
        }

        private GA.RunReportResponse ExecutarConsulta(string property, string start, string end, string? dimension, params string[] metrics)
        {
            var req = new GA.RunReportRequest
            {
                Property = property,
                DateRanges = { new GA.DateRange { StartDate = start, EndDate = end } },
                Limit = 10000 
            };
            
            if (!string.IsNullOrEmpty(dimension))
                req.Dimensions.Add(new GA.Dimension { Name = dimension });

            foreach (var m in metrics)
                req.Metrics.Add(new GA.Metric { Name = m });
            
            if (!string.IsNullOrEmpty(dimension))
            {
                req.OrderBys.Add(new GA.OrderBy
                {
                    Metric = new GA.OrderBy.Types.MetricOrderBy { MetricName = metrics[0] },
                    Desc = true
                });
            }

            return _client.RunReport(req);
        }

        private void PreencherLista(List<string> lista, string property, string start, string end, string dimension, string metric, int limit, double totalReferencia)
        {
            var res = ExecutarConsulta(property, start, end, dimension, metric);
            
            int pos = 1;
            foreach (var row in res.Rows.Take(limit))
            {
                if (row.DimensionValues.Count == 0 || row.MetricValues.Count == 0) continue;

                string nome = row.DimensionValues[0].Value;
                string valorStr = row.MetricValues[0].Value;
                double.TryParse(valorStr, out double valorNum);
                
                double pct = totalReferencia > 0 ? (valorNum / totalReferencia) * 100 : 0;
                
                lista.Add($"{pos} {nome} - {valorStr} ({pct:F2}%)");
                pos++;
            }
        }
    }
    
    public class SlideService
    {
        public void Gerar(Cliente cliente, DadosAnalytics dados, string pastaDestino)
        {
            if (!Directory.Exists(pastaDestino))
                Directory.CreateDirectory(pastaDestino);

            string nomeArquivo = $"Relatorio Mensal de acessos ao site - {dados.Periodo.NomeMes} de {dados.Periodo.Ano} - {cliente.Nome.Replace(" ", "_")} - {cliente.Estado.Replace(" ", "_")}.pptx";
            string caminhoFinal = Path.Combine(pastaDestino, nomeArquivo);

            Console.WriteLine($"\nIniciando geração para {cliente.Nome}...");
            
            Console.WriteLine("Atualizando histórico de gráficos no modelo base...");
            using (PresentationDocument pptTemplate = PresentationDocument.Open(cliente.CaminhoTemplateSlide, true))
            {
                var pPartTemplate = pptTemplate.PresentationPart!;
                var slideIdsTemplate = pPartTemplate.Presentation.SlideIdList!.Elements<P.SlideId>().ToList();

                double.TryParse(dados.TotalSessoes, out double totalSessoes);
                double.TryParse(dados.TotalPageViews, out double totalViews);
                
                if (slideIdsTemplate.Count > 1)
                    AtualizarGrafico(pPartTemplate, slideIdsTemplate[1], dados.Periodo.NumeroMes, new[] { totalSessoes });
                
                if (slideIdsTemplate.Count > 2)
                    AtualizarGrafico(pPartTemplate, slideIdsTemplate[2], dados.Periodo.NumeroMes, new[] { dados.UsuariosDesktop, dados.UsuariosMobile });
                
                if (slideIdsTemplate.Count > 6)
                    AtualizarGrafico(pPartTemplate, slideIdsTemplate[6], dados.Periodo.NumeroMes, new[] { totalViews });
            }
            
            if (File.Exists(caminhoFinal))
            {
                try
                {
                    File.SetAttributes(caminhoFinal, FileAttributes.Normal);
                    File.Delete(caminhoFinal);
                }
                catch (IOException)
                {
                    Console.WriteLine($"[AVISO] O arquivo {nomeArquivo} está ABERTO no PowerPoint!");
                    Console.WriteLine("Feche-o para atualizar. Pulando cliente...");
                    return;
                }
            }
            File.Copy(cliente.CaminhoTemplateSlide, caminhoFinal, true);
            
            Console.WriteLine("Injetando textos e listas no relatório final...");
            using (PresentationDocument pptFinal = PresentationDocument.Open(caminhoFinal, true))
            {
                var pPartFinal = pptFinal.PresentationPart!;
                var slideIdsFinal = pPartFinal.Presentation.SlideIdList!.Elements<P.SlideId>().ToList();
                
                foreach (var slideId in slideIdsFinal)
                {
                    SlidePart slidePart = (SlidePart)pPartFinal.GetPartById(slideId.RelationshipId!.Value!);
                    SubstituirTextos(slidePart, dados);
                    SubstituirListas(slidePart, dados);
                }
            }
            
            Console.WriteLine($"Slide concluído e salvo em: {caminhoFinal}");
        }

        private void SubstituirTextos(SlidePart slidePart, DadosAnalytics dados)
        {
            var dic = new Dictionary<string, string>
            {
                { "#MES_NOME#", dados.Periodo.NomeMes },
                { "#ANO#", dados.Periodo.Ano },
                { "#DIAS_MES#", dados.Periodo.DiasFormatados },
                { "#SESSOES#", dados.TotalSessoes },
                { "#USUARIOS#", dados.TotalUsuarios },
                { "#MOB_PCT#", dados.PctMobile },
                { "#DESK_PCT#", dados.PctDesktop },
                { "#TOTAL_NAVEG#", dados.TotalUsuariosAtivos },
                { "#TOTAL_RESOL#", dados.TotalUsuariosAtivos },
                { "#TOTAL_CIDADES#", dados.TotalUsuariosAtivos },
                { "#TOTAL_PAGINAS#", dados.TotalPageViews }
            };

            foreach (var t in slidePart.Slide.Descendants<A.Text>().Where(x => !string.IsNullOrEmpty(x.Text)))
            {
                foreach (var kvp in dic)
                {
                    if (t.Text.Contains(kvp.Key))
                        t.Text = t.Text.Replace(kvp.Key, kvp.Value);
                }
            }
            slidePart.Slide.Save();
        }

        private void SubstituirListas(SlidePart slidePart, DadosAnalytics dados)
        {
            ProcessarListaTag(slidePart, "#NAVEGADORES#", dados.ListaNavegadores);
            ProcessarListaTag(slidePart, "#RESOLUCOES#", dados.ListaResolucoes);
            ProcessarListaTag(slidePart, "#CIDADES#", dados.ListaCidades);
            ProcessarListaTag(slidePart, "#PAGINAS#", dados.ListaPaginas);
            slidePart.Slide.Save();
        }

        private void ProcessarListaTag(SlidePart slidePart, string tag, List<string> linhas)
        {
            var textoMarcador = slidePart.Slide.Descendants<A.Text>().FirstOrDefault(t => t.Text.Contains(tag));
            if (textoMarcador == null) return;

            var runOriginal = textoMarcador.Parent as A.Run;
            var paragrafo = runOriginal?.Parent as A.Paragraph;
            if (runOriginal == null || paragrafo == null) return;

            if (linhas.Count == 0)
            {
                textoMarcador.Text = "Sem dados no período.";
                return;
            }

            OpenXmlElement ultimoElemento = runOriginal;
            for (int i = 0; i < linhas.Count; i++)
            {
                var novoRun = (A.Run)runOriginal.CloneNode(true);
                novoRun.GetFirstChild<A.Text>()!.Text = linhas[i];
                paragrafo.InsertAfter(novoRun, ultimoElemento);
                ultimoElemento = novoRun;

                if (i < linhas.Count - 1)
                {
                    var quebra = new A.Break();
                    paragrafo.InsertAfter(quebra, ultimoElemento);
                    ultimoElemento = quebra;
                }
            }
            runOriginal.Remove(); 
        }

        private void AtualizarGrafico(PresentationPart pPart, P.SlideId slideId, int mesAtual, double[] novosValores)
        {
            SlidePart slidePart = (SlidePart)pPart.GetPartById(slideId.RelationshipId!.Value!);
            ChartPart? chartPart = slidePart.ChartParts.FirstOrDefault();
            if (chartPart == null) return;
            
            var excelPart = chartPart.EmbeddedPackagePart;
            if (excelPart != null)
            {
                using (Stream stream = excelPart.GetStream(FileMode.Open, FileAccess.ReadWrite))
                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(stream, true))
                {
                    var workbookPart = spreadsheet.WorkbookPart!;
                    var sheet = workbookPart.Workbook.Descendants<S.Sheet>().FirstOrDefault();
                    if (sheet != null)
                    {
                        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
                        var sheetData = worksheetPart.Worksheet.Elements<S.SheetData>().FirstOrDefault();

                        if (sheetData != null)
                        {
                            uint rowIndex = (uint)(mesAtual + 1); 
                            
                            S.Row? row = sheetData.Elements<S.Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIndex);
                            if (row == null)
                            {
                                row = new S.Row() { RowIndex = new UInt32Value(rowIndex) };
                                sheetData.Append(row);
                            }

                            string[] colunas = { "B", "C", "D", "E" };
                            for (int i = 0; i < novosValores.Length && i < colunas.Length; i++)
                            {
                                string cellRef = $"{colunas[i]}{rowIndex}";
                                S.Cell? cell = row.Elements<S.Cell>().FirstOrDefault(c => c.CellReference?.Value == cellRef);
                                if (cell == null)
                                {
                                    cell = new S.Cell() { CellReference = new StringValue(cellRef) };
                                    row.Append(cell);
                                }

                                cell.CellValue = new S.CellValue(novosValores[i].ToString(CultureInfo.InvariantCulture));
                                cell.DataType = new EnumValue<S.CellValues>(S.CellValues.Number);
                            }
                        }
                    }
                }
            }
            
            var chartSpace = chartPart.ChartSpace;
            var seriesList = chartSpace.Descendants<C.SeriesText>().Select(s => s.Parent).ToList();
            uint ptIndex = (uint)(mesAtual - 1); 

            for (int i = 0; i < seriesList.Count && i < novosValores.Length; i++)
            {
                var valReference = seriesList[i]!.Descendants<C.NumberReference>().FirstOrDefault();
                if (valReference?.NumberingCache != null)
                {
                    var pt = valReference.NumberingCache.Descendants<C.NumericPoint>().FirstOrDefault(p => p.Index?.Value == ptIndex);
                    
                    if (pt?.NumericValue != null)
                    {
                        pt.NumericValue.Text = novosValores[i].ToString(CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        var newPt = new C.NumericPoint() { Index = new UInt32Value(ptIndex) };
                        newPt.Append(new C.NumericValue(novosValores[i].ToString(CultureInfo.InvariantCulture)));
                        valReference.NumberingCache.Append(newPt);

                        var ptCount = valReference.NumberingCache.Descendants<C.PointCount>().FirstOrDefault();
                        if (ptCount?.Val != null && ptCount.Val.Value <= ptIndex)
                        {
                            ptCount.Val.Value = ptIndex + 1;
                        }
                    }
                }

                var dLbl = seriesList[i]!.Descendants<C.DataLabel>().FirstOrDefault(l => l.Index?.Val?.Value == ptIndex);
                dLbl?.Descendants<C.ChartText>().FirstOrDefault()?.Remove();
            }
            
            chartPart.ChartSpace.Save();
        }
    }
    
    class Program
    {
        static void Main()
        {
            string caminhoCredencial = @"C:\Users\samu0\Documents\Analytics\analytics-automacao-1c71b3e01156.json";
            string caminhoClientesJson = @"C:\Users\samu0\Documents\Analytics\ClientesdaRIFT.json";
            string pastaDestino = @"C:\Users\samu0\Documents\Analytics\islaide gerado";

            if (!File.Exists(caminhoClientesJson))
            {
                Console.WriteLine("Arquivo JSON de clientes não encontrado no caminho especificado!");
                return;
            }

            var jsonText = File.ReadAllText(caminhoClientesJson);
            var clientes = JsonSerializer.Deserialize<List<Cliente>>(jsonText);

            if (clientes == null || clientes.Count == 0)
            {
                Console.WriteLine("A lista de clientes está vazia ou o JSON está inválido.");
                return;
            }

            var analyticsService = new AnalyticsService(caminhoCredencial);
            var slideService = new SlideService();

            Console.WriteLine("--- Iniciando Automação de Relatórios ---");

            foreach (var cliente in clientes)
            {
                try
                {
                    Console.WriteLine($"\nProcessando: {cliente.Nome}");
                    var dados = analyticsService.ObterDados(cliente.Ga4PropertyId);
                    slideService.Gerar(cliente, dados, pastaDestino);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[ERRO] Falha no Cliente {cliente.Nome}: {ex.Message}");
                }
            }
            Console.WriteLine("\n--- Processo Finalizado com Sucesso! ---");
        }
    }
}