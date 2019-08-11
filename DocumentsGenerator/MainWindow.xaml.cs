using DocumentsGenerator.Models;
using DocumentsGenerator.ViewModels;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Windows;
using System.Windows.Input;

namespace DocumentsGenerator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            var dialogResult = new DialogWindow().ShowDialog();
            if (dialogResult.Value == true)
            {
                BinaryFormatter formatter = new BinaryFormatter();
                using (FileStream fs = new FileStream("data.dat", FileMode.OpenOrCreate))
                {
                    var viewModel = this.DataContext as MainViewModel;
                    var docData = new DocumentData()
                    {
                        ContractId = viewModel.ContractId,
                        ContractDate = viewModel.ContractDate,
                        CompanyName = viewModel.CompanyName,
                        StartRentDate = viewModel.StartRentDate,
                        EndRentDate = viewModel.EndRentDate,
                        PostIndex = viewModel.PostIndex,
                        Adress = viewModel.Adress,
                        FacktureId = viewModel.SettlementAccount,
                        MFO = viewModel.BankMFO,
                        Bank = viewModel.BankName,
                        EDROPOY = viewModel.CompanyYEDROPOU,
                        AccountId = viewModel.AccountId,
                        AccountDate = viewModel.AccountDate,
                        ActId = viewModel.ActId,
                        ActDate = viewModel.ActDate,
                        DirectoryName = viewModel.CompanyDirector,
                        TotalAmount = viewModel.TotalAmount,
                        TotalAmountInWords = viewModel.TotalAmountInWords,
                        AmountWithoutPDV = viewModel.TotalAmountWithoutPDV,
                        AmountWithoutPDVInWords = viewModel.TotalAmountWithoutPDVInWords,
                        EquipmentUsingAdress = viewModel.EquipmentUsingAdress
                    };
                    formatter.Serialize(fs, docData);
                }
            }
            else
            {
                if (File.Exists("data.dat")) File.Delete("data.dat");  
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            BinaryFormatter formatter = new BinaryFormatter();
            using (FileStream fs = new FileStream("data.dat",FileMode.OpenOrCreate))
            {
                try
                {
                    var documentData = (DocumentData)formatter.Deserialize(fs);
                    var viewModel = this.DataContext as MainViewModel;
                    viewModel.ContractId = documentData.ContractId;
                    viewModel.ContractDate = documentData.ContractDate;
                    viewModel.CompanyName = documentData.CompanyName;
                    viewModel.StartRentDate = documentData.StartRentDate;
                    viewModel.EndRentDate = documentData.EndRentDate;
                    viewModel.PostIndex = documentData.PostIndex;
                    viewModel.Adress = documentData.Adress;
                    viewModel.SettlementAccount = documentData.FacktureId;
                    viewModel.BankMFO = documentData.MFO;
                    viewModel.BankName = documentData.Bank;
                    viewModel.CompanyYEDROPOU = documentData.EDROPOY;
                    viewModel.AccountId = documentData.AccountId;
                    viewModel.AccountDate = documentData.AccountDate;
                    viewModel.ActId = documentData.ActId;
                    viewModel.ActDate = documentData.ActDate;
                    viewModel.CompanyDirector = documentData.DirectoryName;
                    viewModel.TotalAmount = documentData.TotalAmount;
                    viewModel.TotalAmountInWords = documentData.TotalAmountInWords;
                    viewModel.TotalAmountWithoutPDV = documentData.AmountWithoutPDV;
                    viewModel.TotalAmountWithoutPDVInWords = documentData.AmountWithoutPDVInWords;
                    viewModel.EquipmentUsingAdress = documentData.EquipmentUsingAdress;
                }
                catch 
                {
                    //ignore
                }
            }
        }

        private void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Down)
            {
                MoveToNextUIElement(e,FocusNavigationDirection.Next);
            }
            if (e.Key == Key.Up)
            {
                MoveToNextUIElement(e, FocusNavigationDirection.Previous);
            }
        }
        void MoveToNextUIElement(KeyEventArgs e, FocusNavigationDirection focusDirection)
        {
            // MoveFocus takes a TraveralReqest as its argument.
            TraversalRequest request = new TraversalRequest(focusDirection);

            // Gets the element with keyboard focus.
            UIElement elementWithFocus = Keyboard.FocusedElement as UIElement;

            // Change keyboard focus.
            if (elementWithFocus != null)
            {
                if (elementWithFocus.MoveFocus(request)) e.Handled = true;
            }
        }
    }
}
