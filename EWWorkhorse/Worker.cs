using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;

namespace EWWorkhorse
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;

        public Worker(ILogger<Worker> logger)
        {
            _logger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                if (_logger.IsEnabled(LogLevel.Information))
                {
                    _logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);
                }
                await Task.Delay(1000, stoppingToken);
                
                // todo: add Replace
            }
        }

        public async Task<WordprocessingDocument> Replace(WordprocessingDocument wdoc, SpreadsheetDocument excDoc)
        {
            var wordBody = wdoc.MainDocumentPart.Document.Body;
            var paragraphs = wordBody.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();
            Regex markerRegEx = new Regex(@"<#\d+#[A-Z]+\d+>");

            foreach (var paragraph in paragraphs)
            {
                foreach (var run in paragraph.Elements<DocumentFormat.OpenXml.Wordprocessing.Run>())
                {
                    foreach (var text in run.Elements<DocumentFormat.OpenXml.Wordprocessing.Text>())
                    {
                        MatchCollection markerMatches = markerRegEx.Matches(text.Text);

                        foreach (Match match in markerMatches)
                        {
                            Regex sheetRegEx = new Regex(@"#\d+#");
                            Regex cellRegEx = new Regex(@"#[A-Z]+\d+>");
                            int sheetIndex = Int32.Parse(sheetRegEx.Match(match.Value).Value.Trim('#'));
                            string cellIndex = cellRegEx.Match(match.Value).Value.Trim('#', '>');
                            WorkbookPart wbPart = excDoc.WorkbookPart;
                            Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.SheetId == sheetIndex);
                            WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
                            Cell cell = wsPart.Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference == cellIndex);

                            var value = cell.InnerText;

                            if (cell.DataType is not null)
                            {
                                if (cell.DataType.Value == CellValues.SharedString)
                                {
                                    var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                                    value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;

                                    text.Text = text.Text.Replace(match.Value, value);
                                }
                            }
                            else
                            {
                                text.Text = text.Text.Replace(match.Value, value);
                            }
                        }
                    }
                }
            }

            wdoc.Save();
            return wdoc;
        }
    }
}
