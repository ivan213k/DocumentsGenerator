using DocumentsGenerator.FOP;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;

namespace DocumentsGenerator
{
    /// <summary>
    /// Interaction logic for ChooseFOPDialogWindow.xaml
    /// </summary>
    public partial class ChooseFOPDialogWindow : Window
    {
        private IList<FopModel> Fops { get; set; }

        public ChooseFOPDialogWindow()
        {
            InitializeComponent();
            InitFops();
            InitializeFOPSelector();
        }

        private void InitializeFOPSelector()
        {
            foreach (var fop in Fops)
            {
                FopsComboBox.Items.Add(fop.Name);
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            FopsInfo.SelectedFop = Fops.Single(f => f.Name == FopsComboBox.Text);
        }

        private void Ok_Button_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void InitFops()
        {
            string fopsJson;
            using (StreamReader streamReader = new StreamReader("FopsInfo.json"))
            {
                fopsJson = streamReader.ReadToEnd();
            }
            Fops = JsonSerializer.Deserialize<List<FopModel>>(fopsJson);
        }
    }
}
