using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace DocumentsGenerator.Models
{
    class ContractGenerator
    {
        private readonly string templateFile = "Договор_Template.doc"; 

        public async Task GenerateContractAsync(List<KeyValuePair<string,string>> valuePairs, string filePath)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(templateFile,true))
            {
                var document = doc.MainDocumentPart.Document;

                foreach (var text in document.Descendants<Text>())
                {
                    if (text.Text.Contains("text-to-replace"))
                    {
                        text.Text = text.Text.Replace("text-to-replace", "replaced-text");
                    }
                }
            }
        }
    }
}
