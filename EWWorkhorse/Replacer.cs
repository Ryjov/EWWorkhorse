using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Google.Protobuf;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace EWWorkhorse
{
    internal static class Replacer
    {
        public static WordprocessingDocument ReplaceFile(byte[] wordBytes, byte[] excBytes)
        {
            using (MemoryStream excelMeme = new MemoryStream())
            using (MemoryStream wordMeme = new MemoryStream())
            {
                excelMeme.Write(excBytes, 0, (int)excBytes.Length);
                wordMeme.Write(wordBytes, 0, (int)wordBytes.Length);
                using (SpreadsheetDocument excDoc = SpreadsheetDocument.Open(excelMeme, true))
                using (WordprocessingDocument wdoc = WordprocessingDocument.Open(wordMeme, true))
                {
                    var result = ReplaceFile(wdoc, excDoc);

                    return result;
                }
            }
        }

        public static WordprocessingDocument ReplaceFile(WordprocessingDocument wdoc, SpreadsheetDocument excDoc)
        {
            var document = wdoc.MainDocumentPart.Document;
            Regex markerRegEx = new Regex(@"<#\d+#[A-Z]+\d+>");

            foreach (var text in document.Descendants<Text>())
            {
                foreach (Match match in markerRegEx.Matches(text.Text))
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

            wdoc.Save();
            return wdoc;
        }
    }
}