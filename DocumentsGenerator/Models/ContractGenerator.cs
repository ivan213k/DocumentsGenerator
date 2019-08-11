using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;

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
                FillTable(null, generatedDocument);
            }
        }

        void FillTable(List<Equipment> equipments, WordprocessingDocument wordDoc)
        {
            var body = wordDoc.MainDocumentPart.Document.Body;
            //var tables = body.Descendants<Table>().Where(tbl => tbl.GetFirstChild<TableRow>().Descendants<TableCell>().Count() == 8);
            foreach (Table table in body.Descendants<Table>().Where(tbl => tbl.GetFirstChild<TableRow>().Descendants<TableCell>().Count() == 8))
            {
                RunProperties runProperties = new RunProperties();
                runProperties.AppendChild(new FontSize() { Val = "20" });
                var run = new Run(new Text(" 1 d ddd"));
                run.PrependChild<RunProperties>(runProperties);
                var lastTableRow = table.Elements<TableRow>().Last();
                var newRow = lastTableRow.CloneNode(true);
                newRow.Descendants<TableCell>().ElementAt(2).Descendants<Paragraph>().First().Descendants<Run>().First().Append(new Text("test"));
                MessageBox.Show(newRow.InnerXml);
               // newRow.Descendants<TableCell>().ElementAt(2).RemoveAllChildren<Paragraph>();
                //newRow.Descendants<TableCell>().ElementAt(2).Append(new Paragraph(run));
                table.Append(newRow);
                //table.Insert(new TableRow(new TableCell(new Paragraph(new Run(new Text("")))), new TableCell(new Paragraph(new Run(new Text("test"))))), 2);
                //table.Append(new TableRow(new TableCell(new Paragraph(new Run(new Text("")))), new TableCell(new Paragraph(new Run(new Text("test"))))));
            }
            wordDoc.Save();
        }
    }
}
