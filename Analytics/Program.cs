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
    // ══════════════════════════════════════════════════════════════════════════
    //  Modelos de Dados
    // ══════════════════════════════════════════════════════════════════════════

    public class Cliente
    {
        public string Nome { get; set; } = "";
        public string Estado { get; set; } = "";
        public string Ga4PropertyId { get; set; } = "";

        /// <summary>
        /// Caminho da pasta do cliente, que contém a logo (logo.png) e o JSON de
        /// histórico anual (historico_YYYY.json).
        /// Ex.: C:\...\Analytics\clientes\CartorioFulano
        /// </summary>
        public string CaminhoClientePasta { get; set; } = "";

        /// <summary>Caminho para o arquivo logo.png do cliente (derivado da pasta).</summary>
        [JsonIgnore]
        public string CaminhoLogo => Path.Combine(CaminhoClientePasta, "logo.png");

        /// <summary>Caminho para o JSON de histórico do ano informado (ex.: historico_2026.json).</summary>
        public string CaminhoHistoricoJson(string ano) =>
            Path.Combine(CaminhoClientePasta, $"historico_{ano}.json");
    }

    /// <summary>Dados de um único mês armazenados no JSON de histórico anual.</summary>
    public class DadosMes
    {
        public string NomeMes  { get; set; } = "";   // Ex.: "JULHO"
        public double Sessoes  { get; set; }
        public double Desktop  { get; set; }
        public double Mobile   { get; set; }
        public double PageViews { get; set; }
    }

    /// <summary>Estrutura completa do JSON de histórico anual por cliente.</summary>
    public class HistoricoAnual
    {
        public string Ano { get; set; } = "";

        /// <summary>Chave: número do mês ("1"–"12"). Valor: dados daquele mês.</summary>
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

    // ══════════════════════════════════════════════════════════════════════════
    //  Serviço Google Analytics
    // ══════════════════════════════════════════════════════════════════════════

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
            string dtFim    = dados.Periodo.DataFim.ToString("yyyy-MM-dd");
            string prop     = $"properties/{propertyId}";

            Console.WriteLine($"Buscando dados de {dtInicio} a {dtFim}...");

            var resGeral = ExecutarConsulta(prop, dtInicio, dtFim, null, "sessions", "totalUsers", "activeUsers", "screenPageViews");
            if (resGeral.Rows.Count > 0)
            {
                dados.TotalSessoes        = resGeral.Rows[0].MetricValues[0].Value;
                dados.TotalUsuarios       = resGeral.Rows[0].MetricValues[1].Value;
                dados.TotalUsuariosAtivos = resGeral.Rows[0].MetricValues[2].Value;
                dados.TotalPageViews      = resGeral.Rows[0].MetricValues[3].Value;
            }

            double.TryParse(dados.TotalSessoes,        out double totalSessoesNum);
            double.TryParse(dados.TotalUsuariosAtivos,  out double totalUsuariosNum);
            double.TryParse(dados.TotalPageViews,       out double totalViewsNum);

            var resDisp = ExecutarConsulta(prop, dtInicio, dtFim, "deviceCategory", "activeUsers");
            foreach (var row in resDisp.Rows)
            {
                if (row.DimensionValues.Count == 0 || row.MetricValues.Count == 0) continue;

                string disp = row.DimensionValues[0].Value.ToLower();
                double.TryParse(row.MetricValues[0].Value, out double val);

                if (disp == "desktop") dados.UsuariosDesktop += val;
                else                   dados.UsuariosMobile  += val;
            }

            double totalDisp = dados.UsuariosDesktop + dados.UsuariosMobile;
            if (totalDisp > 0)
            {
                dados.PctDesktop = $"{(dados.UsuariosDesktop / totalDisp) * 100:F2}%";
                dados.PctMobile  = $"{(dados.UsuariosMobile  / totalDisp) * 100:F2}%";
            }

            PreencherLista(dados.ListaNavegadores, prop, dtInicio, dtFim, "browser",           "activeUsers",    7,  totalUsuariosNum);
            PreencherLista(dados.ListaResolucoes,  prop, dtInicio, dtFim, "screenResolution",  "activeUsers",    10, totalUsuariosNum);
            PreencherLista(dados.ListaCidades,     prop, dtInicio, dtFim, "city",              "activeUsers",    10, totalUsuariosNum);
            PreencherLista(dados.ListaPaginas,     prop, dtInicio, dtFim, "pageTitle",         "screenPageViews",10, totalViewsNum);

            return dados;
        }

        private GA.RunReportResponse ExecutarConsulta(string property, string start, string end,
                                                      string? dimension, params string[] metrics)
        {
            var req = new GA.RunReportRequest
            {
                Property   = property,
                DateRanges = { new GA.DateRange { StartDate = start, EndDate = end } },
                Limit      = 10000
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
                    Desc   = true
                });
            }

            return _client.RunReport(req);
        }

        private void PreencherLista(List<string> lista, string property, string start, string end,
                                    string dimension, string metric, int limit, double totalReferencia)
        {
            var res = ExecutarConsulta(property, start, end, dimension, metric);

            int pos = 1;
            foreach (var row in res.Rows.Take(limit))
            {
                if (row.DimensionValues.Count == 0 || row.MetricValues.Count == 0) continue;

                string nome     = row.DimensionValues[0].Value;
                string valorStr = row.MetricValues[0].Value;
                double.TryParse(valorStr, out double valorNum);

                double pct = totalReferencia > 0 ? (valorNum / totalReferencia) * 100 : 0;

                lista.Add($"{pos} {nome} - {valorStr} ({pct:F2}%)");
                pos++;
            }
        }
    }

    // ══════════════════════════════════════════════════════════════════════════
    //  Serviço de Histórico Anual (JSON por cliente)
    // ══════════════════════════════════════════════════════════════════════════

    public class HistoricoService
    {
        private static readonly JsonSerializerOptions JsonOpts = new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder       = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        /// <summary>
        /// Lê o JSON de histórico do caminho indicado. Se não existir ou estiver corrompido,
        /// retorna um <see cref="HistoricoAnual"/> vazio para o ano informado.
        /// </summary>
        public HistoricoAnual CarregarOuCriar(string caminhoJson, string ano)
        {
            if (File.Exists(caminhoJson))
            {
                try
                {
                    var texto     = File.ReadAllText(caminhoJson);
                    var historico = JsonSerializer.Deserialize<HistoricoAnual>(texto, JsonOpts);
                    if (historico != null) return historico;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[AVISO] Erro ao ler histórico JSON: {ex.Message}. Criando novo.");
                }
            }

            return new HistoricoAnual { Ano = ano, Meses = new Dictionary<string, DadosMes>() };
        }

        /// <summary>
        /// Insere ou atualiza no objeto <paramref name="historico"/> os dados do mês atual e
        /// persiste o JSON no disco.
        /// </summary>
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

            string? pasta = Path.GetDirectoryName(caminhoJson);
            if (!string.IsNullOrEmpty(pasta) && !Directory.Exists(pasta))
                Directory.CreateDirectory(pasta);

            File.WriteAllText(caminhoJson, JsonSerializer.Serialize(historico, JsonOpts));
            Console.WriteLine($"Histórico atualizado: {caminhoJson}");
        }
    }

    // ══════════════════════════════════════════════════════════════════════════
    //  Serviço de Geração de Slides
    // ══════════════════════════════════════════════════════════════════════════

    public class SlideService
    {
        private readonly HistoricoService _historicoService = new HistoricoService();

        /// <summary>
        /// Fluxo completo de geração de um slide para um cliente:
        /// 1. Garante pasta do cliente | 2. Atualiza histórico JSON | 3. Copia template |
        /// 4. Popula gráficos com histórico completo + escala dinâmica |
        /// 5. Injeta textos/listas | 6. Substitui logo.
        /// O template original NUNCA é modificado.
        /// </summary>
        public void Gerar(Cliente cliente, DadosAnalytics dados, string caminhoTemplate, string pastaDestino)
        {
            // ── 1. Garantir pasta do cliente ──────────────────────────────────
            if (!Directory.Exists(cliente.CaminhoClientePasta))
                Directory.CreateDirectory(cliente.CaminhoClientePasta);

            // ── 2. Carregar histórico, atualizar mês atual e salvar ───────────
            string caminhoHistorico = cliente.CaminhoHistoricoJson(dados.Periodo.Ano);
            var historico = _historicoService.CarregarOuCriar(caminhoHistorico, dados.Periodo.Ano);
            _historicoService.SalvarMesAtual(historico, dados, caminhoHistorico);

            // ── 3. Montar caminho do arquivo de saída ─────────────────────────
            if (!Directory.Exists(pastaDestino))
                Directory.CreateDirectory(pastaDestino);

            string nomeArquivo = $"Relatorio Mensal de acessos ao site - {dados.Periodo.NomeMes} de {dados.Periodo.Ano}" +
                                 $" - {cliente.Nome.Replace(" ", "_")} - {cliente.Estado.Replace(" ", "_")}.pptx";
            string caminhoFinal = Path.Combine(pastaDestino, nomeArquivo);

            Console.WriteLine($"\nIniciando geração para {cliente.Nome}...");

            // ── 4. Verificar se o arquivo de saída está em uso ────────────────
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

            // ── 5. Copiar template → destino (template original intocado) ─────
            File.Copy(caminhoTemplate, caminhoFinal, true);

            // ── 6. Abrir CÓPIA e aplicar todas as modificações ────────────────
            Console.WriteLine("Atualizando gráficos com histórico completo...");
            using (PresentationDocument ppt = PresentationDocument.Open(caminhoFinal, true))
            {
                var pPart    = ppt.PresentationPart!;
                var slideIds = pPart.Presentation.SlideIdList!.Elements<P.SlideId>().ToList();

                // Slide índice 1 — Sessões (1 série)
                if (slideIds.Count > 1)
                {
                    var valores = ExtrairValores(historico, m => new[] { m.Sessoes });
                    AtualizarGrafico(pPart, slideIds[1], valores);
                }

                // Slide índice 2 — Dispositivos: Desktop + Mobile (2 séries)
                if (slideIds.Count > 2)
                {
                    var valores = ExtrairValores(historico, m => new[] { m.Desktop, m.Mobile });
                    AtualizarGrafico(pPart, slideIds[2], valores);
                }

                // Slide índice 6 — PageViews (1 série)
                if (slideIds.Count > 6)
                {
                    var valores = ExtrairValores(historico, m => new[] { m.PageViews });
                    AtualizarGrafico(pPart, slideIds[6], valores);
                }

                Console.WriteLine("Injetando textos e listas...");
                foreach (var slideId in slideIds)
                {
                    var slidePart = (SlidePart)pPart.GetPartById(slideId.RelationshipId!.Value!);
                    SubstituirTextos(slidePart, dados);
                    SubstituirListas(slidePart, dados);
                }

                Console.WriteLine("Substituindo logo...");
                foreach (var slideId in slideIds)
                {
                    var slidePart = (SlidePart)pPart.GetPartById(slideId.RelationshipId!.Value!);
                    SubstituirLogo(slidePart, cliente.CaminhoLogo);
                }
            }

            Console.WriteLine($"Slide concluído e salvo em: {caminhoFinal}");
        }

        // ── Extrai lista ordenada de (Mes, Valores[]) a partir do histórico ───
        private List<(int Mes, double[] Valores)> ExtrairValores(
            HistoricoAnual historico, Func<DadosMes, double[]> seletor)
        {
            return historico.Meses
                .Select(kv => (Mes: int.Parse(kv.Key), Valores: seletor(kv.Value)))
                .OrderBy(x => x.Mes)
                .ToList();
        }

        // ─────────────────────────────────────────────────────────────────────
        //  Substituição de Textos
        // ─────────────────────────────────────────────────────────────────────

        private void SubstituirTextos(SlidePart slidePart, DadosAnalytics dados)
        {
            var dic = new Dictionary<string, string>
            {
                { "#MES_NOME#",      dados.Periodo.NomeMes         },
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

        // ─────────────────────────────────────────────────────────────────────
        //  Substituição de Listas
        // ─────────────────────────────────────────────────────────────────────

        private void SubstituirListas(SlidePart slidePart, DadosAnalytics dados)
        {
            ProcessarListaTag(slidePart, "#NAVEGADORES#", dados.ListaNavegadores);
            ProcessarListaTag(slidePart, "#RESOLUCOES#",  dados.ListaResolucoes);
            ProcessarListaTag(slidePart, "#CIDADES#",     dados.ListaCidades);
            ProcessarListaTag(slidePart, "#PAGINAS#",     dados.ListaPaginas);
            slidePart.Slide.Save();
        }

        private void ProcessarListaTag(SlidePart slidePart, string tag, List<string> linhas)
        {
            var textoMarcador = slidePart.Slide.Descendants<A.Text>()
                .FirstOrDefault(t => t.Text.Contains(tag));
            if (textoMarcador == null) return;

            var runOriginal = textoMarcador.Parent as A.Run;
            var paragrafo   = runOriginal?.Parent as A.Paragraph;
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

        // ─────────────────────────────────────────────────────────────────────
        //  Substituição de Logo (Picture Shape nomeado "#LOGO#")
        // ─────────────────────────────────────────────────────────────────────

        private void SubstituirLogo(SlidePart slidePart, string caminhoLogo)
        {
            if (!File.Exists(caminhoLogo))
            {
                Console.WriteLine($"[AVISO] Logo não encontrada: {caminhoLogo}. Pulando substituição.");
                return;
            }

            // Localizar picture shape cujo atributo name seja "#LOGO#"
            P.Picture? logoPic = null;
            foreach (var pic in slidePart.Slide.Descendants<P.Picture>())
            {
                var cNvPr = pic.NonVisualPictureProperties?.NonVisualDrawingProperties;
                if (cNvPr?.Name?.Value == "#LOGO#")
                {
                    logoPic = pic;
                    break;
                }
            }

            // Shape "#LOGO#" ausente neste slide — comportamento normal em slides sem logo
            if (logoPic == null) return;

            // Determinar content-type pela extensão do arquivo de logo
            string ext = Path.GetExtension(caminhoLogo).ToLowerInvariant();
            string contentType = ext switch
            {
                ".png"  => "image/png",
                ".jpg"  => "image/jpeg",
                ".jpeg" => "image/jpeg",
                ".gif"  => "image/gif",
                ".bmp"  => "image/bmp",
                ".webp" => "image/webp",
                _       => "image/png"
            };

            // Adicionar nova ImagePart ao SlidePart e alimentá-la com o arquivo da logo
            ImagePart imagePart = slidePart.AddImagePart(contentType);
            using (var stream = File.OpenRead(caminhoLogo))
                imagePart.FeedData(stream);

            string novoRelId = slidePart.GetIdOfPart(imagePart);

            // Atualizar o r:embed do <a:blip> para apontar para a nova imagem
            var blip = logoPic.BlipFill?.Blip;
            if (blip != null)
                blip.Embed = novoRelId;

            slidePart.Slide.Save();
        }

        // ─────────────────────────────────────────────────────────────────────
        //  Atualização de Gráfico (Excel embutido + cache + escala do eixo)
        // ─────────────────────────────────────────────────────────────────────

        private void AtualizarGrafico(PresentationPart pPart, P.SlideId slideId,
                                      List<(int Mes, double[] Valores)> dadosPorMes)
        {
            SlidePart slidePart  = (SlidePart)pPart.GetPartById(slideId.RelationshipId!.Value!);
            ChartPart? chartPart = slidePart.ChartParts.FirstOrDefault();
            if (chartPart == null) return;

            // ── Atualizar planilha Excel embutida ─────────────────────────────
            var excelPart = chartPart.EmbeddedPackagePart;
            if (excelPart != null)
            {
                using Stream stream            = excelPart.GetStream(FileMode.Open, FileAccess.ReadWrite);
                using SpreadsheetDocument xlsx = SpreadsheetDocument.Open(stream, true);

                var workbookPart  = xlsx.WorkbookPart!;
                var sheet         = workbookPart.Workbook.Descendants<S.Sheet>().FirstOrDefault();
                if (sheet != null)
                {
                    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
                    var sheetData     = worksheetPart.Worksheet.Elements<S.SheetData>().FirstOrDefault();

                    if (sheetData != null)
                    {
                        foreach (var (mes, valores) in dadosPorMes)
                        {
                            uint rowIndex = (uint)(mes + 1);  // linha 1 = cabeçalho; mês 1 → linha 2
                            S.Row? row = sheetData.Elements<S.Row>()
                                .FirstOrDefault(r => r.RowIndex?.Value == rowIndex);
                            if (row == null)
                            {
                                row = new S.Row { RowIndex = new UInt32Value(rowIndex) };
                                sheetData.Append(row);
                            }

                            string[] colunas = { "B", "C", "D", "E" };
                            for (int i = 0; i < valores.Length && i < colunas.Length; i++)
                            {
                                string cellRef = $"{colunas[i]}{rowIndex}";
                                S.Cell? cell = row.Elements<S.Cell>()
                                    .FirstOrDefault(c => c.CellReference?.Value == cellRef);
                                if (cell == null)
                                {
                                    cell = new S.Cell { CellReference = new StringValue(cellRef) };
                                    row.Append(cell);
                                }
                                cell.CellValue = new S.CellValue(valores[i].ToString(CultureInfo.InvariantCulture));
                                cell.DataType  = new EnumValue<S.CellValues>(S.CellValues.Number);
                            }
                        }
                    }
                }
            }

            // ── Atualizar cache do gráfico (NumberingCache) ───────────────────
            var chartSpace = chartPart.ChartSpace;
            var seriesList = chartSpace.Descendants<C.SeriesText>().Select(s => s.Parent).ToList();

            foreach (var (mes, valores) in dadosPorMes)
            {
                uint ptIndex = (uint)(mes - 1);  // mês 1 → índice 0

                for (int i = 0; i < seriesList.Count && i < valores.Length; i++)
                {
                    var valRef = seriesList[i]!.Descendants<C.NumberReference>().FirstOrDefault();
                    if (valRef?.NumberingCache == null) continue;

                    var pt = valRef.NumberingCache.Descendants<C.NumericPoint>()
                        .FirstOrDefault(p => p.Index?.Value == ptIndex);

                    if (pt?.NumericValue != null)
                    {
                        pt.NumericValue.Text = valores[i].ToString(CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        var newPt = new C.NumericPoint { Index = new UInt32Value(ptIndex) };
                        newPt.Append(new C.NumericValue(valores[i].ToString(CultureInfo.InvariantCulture)));
                        valRef.NumberingCache.Append(newPt);

                        var ptCount = valRef.NumberingCache.Descendants<C.PointCount>().FirstOrDefault();
                        if (ptCount?.Val != null && ptCount.Val.Value <= ptIndex)
                            ptCount.Val.Value = ptIndex + 1;
                    }

                    // Remover rótulo customizado existente para não conflitar com o novo valor
                    var dLbl = seriesList[i]!.Descendants<C.DataLabel>()
                        .FirstOrDefault(l => l.Index?.Val?.Value == ptIndex);
                    dLbl?.Descendants<C.ChartText>().FirstOrDefault()?.Remove();
                }
            }

            // ── Escala dinâmica do eixo vertical ──────────────────────────────
            double maxValorDados = dadosPorMes
                .SelectMany(x => x.Valores)
                .DefaultIfEmpty(0)
                .Max();

            var (axisMax, majorUnit) = CalcularEscalaEixo(maxValorDados);

            var valAx = chartSpace.Descendants<C.ValueAxis>().FirstOrDefault();
            if (valAx != null)
            {
                // Máximo do eixo
                var scaling = valAx.Descendants<C.Scaling>().FirstOrDefault();
                if (scaling != null)
                {
                    var maxElem = scaling.Elements<C.MaxAxisValue>().FirstOrDefault();
                    if (maxElem == null)
                    {
                        maxElem = new C.MaxAxisValue();
                        scaling.Append(maxElem);
                    }
                    maxElem.Val = axisMax;
                }

                // Unidade principal — controla linhas horizontais visíveis e labels do eixo
                var majorElem = valAx.Elements<C.MajorUnit>().FirstOrDefault();
                if (majorElem == null)
                {
                    majorElem = new C.MajorUnit();
                    valAx.Append(majorElem);
                }
                majorElem.Val = majorUnit;

                // Unidade secundária (subdivisão interna = majorUnit / 5)
                var minorElem = valAx.Elements<C.MinorUnit>().FirstOrDefault();
                if (minorElem == null)
                {
                    minorElem = new C.MinorUnit();
                    valAx.Append(minorElem);
                }
                minorElem.Val = majorUnit / 5.0;
            }

            chartPart.ChartSpace.Save();
        }

        // ─────────────────────────────────────────────────────────────────────
        //  Cálculo de Escala do Eixo Vertical
        //  Regras: máximo ≥ valorMax × 1.25 | exatamente 9 intervalos | unidade "bonita"
        // ─────────────────────────────────────────────────────────────────────

        /// <summary>
        /// Calcula o máximo do eixo e a unidade principal para que o gráfico exiba
        /// exatamente 9 linhas horizontais entre 0 e o máximo, com o topo pelo menos
        /// 25% acima do maior valor presente nos dados.
        /// </summary>
        private (double AxisMax, double MajorUnit) CalcularEscalaEixo(double maxValorDados)
        {
            if (maxValorDados <= 0) maxValorDados = 9;  // mínimo para evitar divisão por zero

            double rawMax  = maxValorDados * 1.25;   // pelo menos 25% acima do máximo dos dados
            double rawUnit = rawMax / 9.0;           // 9 intervalos = 9 linhas horizontais
            double major   = ArredondarNiceNumber(rawUnit);  // arredonda para número legível
            double axisMax = major * 9.0;            // recalcula máximo com a unidade arredondada

            return (axisMax, major);
        }

        /// <summary>
        /// Arredonda <paramref name="valor"/> para cima para o número "bonito" mais próximo
        /// da sequência 1 × 10^n, 2 × 10^n, 2.5 × 10^n, 5 × 10^n, 10 × 10^n.
        /// Exemplos: 49,6 → 50 | 4.167 → 5.000 | 11,1 → 20 | 29,3 → 50
        /// </summary>
        private double ArredondarNiceNumber(double valor)
        {
            if (valor <= 0) return 1;

            double expoente = Math.Floor(Math.Log10(valor));
            double base10   = Math.Pow(10, expoente);
            double fracao   = valor / base10;

            double nice;
            if      (fracao <= 1.0) nice = 1.0;
            else if (fracao <= 2.0) nice = 2.0;
            else if (fracao <= 2.5) nice = 2.5;
            else if (fracao <= 5.0) nice = 5.0;
            else                    nice = 10.0;

            return nice * base10;
        }
    }

    // ══════════════════════════════════════════════════════════════════════════
    //  Ponto de Entrada
    // ══════════════════════════════════════════════════════════════════════════

    class Program
    {
        static void Main()
        {
            // ── Caminhos de configuração (editar conforme necessário) ──────────
            string caminhoCredencial   = @"C:\Users\samu0\Documents\Analytics\analytics-automacao-1c71b3e01156.json";
            string caminhoClientesJson = @"C:\Users\samu0\Documents\Analytics\ClientesdaRIFT.json";
            string caminhoTemplate     = @"C:\Users\samu0\Documents\Analytics\template.pptx";
            string pastaDestino        = @"C:\Users\samu0\Documents\Analytics\islaide gerado";

            // ── Validações iniciais ───────────────────────────────────────────
            if (!File.Exists(caminhoClientesJson))
            {
                Console.WriteLine($"[ERRO] Arquivo de clientes não encontrado: {caminhoClientesJson}");
                return;
            }

            if (!File.Exists(caminhoTemplate))
            {
                Console.WriteLine($"[ERRO] Template não encontrado: {caminhoTemplate}");
                return;
            }

            var jsonText = File.ReadAllText(caminhoClientesJson);
            var clientes = JsonSerializer.Deserialize<List<Cliente>>(jsonText);

            if (clientes == null || clientes.Count == 0)
            {
                Console.WriteLine("[ERRO] A lista de clientes está vazia ou o JSON está inválido.");
                return;
            }

            // ── Processamento ─────────────────────────────────────────────────
            var analyticsService = new AnalyticsService(caminhoCredencial);
            var slideService     = new SlideService();

            Console.WriteLine("--- Iniciando Automação de Relatórios ---");

            foreach (var cliente in clientes)
            {
                try
                {
                    Console.WriteLine($"\nProcessando: {cliente.Nome}");
                    var dados = analyticsService.ObterDados(cliente.Ga4PropertyId);
                    slideService.Gerar(cliente, dados, caminhoTemplate, pastaDestino);
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