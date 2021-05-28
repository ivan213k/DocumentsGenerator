using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DocumentsGenerator.Models
{
    class ContractGenerator
    {
        private readonly string templateFile = "Договор_Template.docx"; 

        public async Task GenerateContract(Dictionary<string,string> valuePairs, IEnumerable<Equipment> equipments, string filePath)
        {
            await Task.Factory.StartNew(()=> 
            {
                string docText = "";
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(templateFile, true))
                {
                    using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                    {
                        docText = sr.ReadToEnd();
                    }

                    docText = ReplaceMarkersByValues(valuePairs, docText);

                    var savedDoc = wordDoc.SaveAs(filePath);
                    savedDoc.Close();
                }
                using (WordprocessingDocument generatedDocument = WordprocessingDocument.Open(filePath, true))
                {
                    using (StreamWriter sw = new StreamWriter(generatedDocument.MainDocumentPart.GetStream(FileMode.Create)))
                    {
                        sw.Write(docText);
                    }
                    FillTable(equipments, generatedDocument);
                    generatedDocument.Save();
                }
            });     
        }

        void FillTable(IEnumerable<Equipment> equipments, WordprocessingDocument wordDoc)
        {
            var table =  wordDoc.MainDocumentPart.Document.Body.Elements<Table>().ToList()[1];
            var firstRow = table.Elements<TableRow>().First();
            var templateRow = table.Elements<TableRow>().Last();
            var newRows = new List<TableRow>();

            int rowIndex = 1;
            foreach (var equipment in equipments)
            {
                var newRow = new TableRow(templateRow.OuterXml);
                var cells = newRow.Elements<TableCell>().ToList();
                cells[0].Elements<Paragraph>().First().Elements<Run>().First().Elements<Text>().First().Text = rowIndex.ToString();
                foreach (var item in equipment.GetValues())
                {
                    cells[item.Key].Elements<Paragraph>().First().Elements<Run>().First().Elements<Text>().First().Text = item.Value;
                }
                newRows.Add(newRow);
                rowIndex++;
            }

            foreach (var row in newRows.Reverse<TableRow>())
            {
                table.InsertAfter(row, firstRow);
            }
        }

        string ReplaceMarkersByValues(Dictionary<string,string> markerValuePairs,string docText)
        {
            foreach (var item in markerValuePairs)
            {
                
                Regex regexText = new Regex(item.Key);
                docText = regexText.Replace(docText, item.Value);
            }
            return docText;
        }
    }
}
