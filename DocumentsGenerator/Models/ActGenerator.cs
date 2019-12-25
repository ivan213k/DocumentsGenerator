using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DocumentsGenerator.Models
{
    class ActGenerator
    {
        private readonly string templateFile = "Акт_Template.xlsx";

        public async Task GenerateAct(Dictionary<string,string> markerValuePairs, string filePath)
        {
            await Task.Factory.StartNew(()=> 
            {
                using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(templateFile, true))
                {
                    var savedDoc = excelDoc.SaveAs(filePath);
                    savedDoc.Close();
                }

                using (SpreadsheetDocument generatedDoc = SpreadsheetDocument.Open(filePath, true))
                {
                    var workTable = generatedDoc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First().SharedStringTable;

                    var innerXml = ReplaceMarkersByValues(markerValuePairs, workTable.InnerXml);
                    workTable.InnerXml = innerXml;

                    workTable.Save();
                }
            });
        }

        string ReplaceMarkersByValues(Dictionary<string, string> markerValuePairs, string docText)
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
