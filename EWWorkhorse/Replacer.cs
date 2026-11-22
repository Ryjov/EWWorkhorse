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
        public static byte[] ReplaceBytes(byte[] wordBytes, byte[] excBytes)
        {
            using var wordMeme = new MemoryStream();
            using var excelMeme = new MemoryStream();
            wordMeme.Write(wordBytes, 0, wordBytes.Length);
            excelMeme.Write(excBytes, 0, excBytes.Length);

            using (var wdoc = WordprocessingDocument.Open(wordMeme, true))
            using (var excDoc = SpreadsheetDocument.Open(excelMeme, true))
            {
                ReplaceFile(wdoc, excDoc);
            }

            return wordMeme.ToArray();
        }

        public static WordprocessingDocument ReplaceFile(WordprocessingDocument wDoc, SpreadsheetDocument excDoc)
        {
            var document = wDoc.MainDocumentPart.Document;
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

            wDoc.Save();
            return wDoc;
        }
    }
}