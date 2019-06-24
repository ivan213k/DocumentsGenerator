using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace DocumentsGenerator.Models
{
    class ContractGenerator
    {
        private readonly string templateFile = "Договор_Template.docx"; 

        public void GenerateContract(Dictionary<string,string> valuePairs, string filePath)
        {
            string docText = "";
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(templateFile,true))
            {
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                foreach (var item in valuePairs)
                {
                    Regex regexText = new Regex(item.Key);
                    docText = regexText.Replace(docText, item.Value);
                }
                
                var savedDoc = wordDoc.SaveAs(filePath);
                savedDoc.Close();
            }
            using (WordprocessingDocument generatedDocument = WordprocessingDocument.Open(filePath, true))
            {
                using (StreamWriter sw = new StreamWriter(generatedDocument.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
            }
        }

    }
}
