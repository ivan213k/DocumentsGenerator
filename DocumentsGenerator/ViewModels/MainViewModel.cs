﻿using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Input;
using DocumentsGenerator.FOP;
using DocumentsGenerator.Models;

namespace DocumentsGenerator.ViewModels
{
    class MainViewModel : BaseViewModel
    {
        private string contractId = "";
        private DateTime? contractDate;
        private string companyName = "";
        private DateTime? startRentDate;
        private DateTime? endRentDate;
        private string postIndex = "";
        private string adress = "";
        private string settlementAccount = "";
        private string bankMFO = "";
        private string bankName = "";
        private string companyYEDROPOU = "";
        private string accountId = "";
        private DateTime? accountDate;
        private string actId = "";
        private DateTime? actDate;
        private string companyDirector = "";
        private decimal totalAmount;
        private string totalAmountInWords = "";
        private decimal totalAmountWithoutPDV;
        private string totalAmountWithoutPDVInWords = "";
        private string equipmentUsingAdress = "";
        MoneyToStrConverter moneyToStrConverter = new MoneyToStrConverter("UAH", "UKR", "F");
        IFormatProvider culture = new CultureInfo("uk-UA");
        public MainViewModel()
        {
            Equipments = new ObservableCollection<Equipment>();
            GenerateDocumentCommand = new Command(GenerateDocuments);
            ClearWindowCommand = new Command(ClearWindow);
            AddEquipmentCommand = new Command(AddEquipment);
            RemoveEquipmentCommand = new Command(RemoveEquipment);
            AddDefultEquipment();
        }

        public string ContractId { get => contractId; set { contractId = value; OnePropertyChanged(); } }

        public DateTime? ContractDate { get => contractDate; set { contractDate = value; OnePropertyChanged(); } }

        public string CompanyName { get => companyName; set { companyName = value; OnePropertyChanged(); } }

        public DateTime? StartRentDate { get => startRentDate; set { startRentDate = value; OnePropertyChanged(); } }

        public DateTime? EndRentDate { get => endRentDate; set { endRentDate = value; OnePropertyChanged(); } }

        public string PostIndex { get => postIndex; set { postIndex = value; OnePropertyChanged(); } }

        public string Adress { get => adress; set { adress = value; OnePropertyChanged(); } }

        public string SettlementAccount { get => settlementAccount; set { settlementAccount = value; OnePropertyChanged(); } }

        public string BankMFO { get => bankMFO; set { bankMFO = value; OnePropertyChanged(); } }

        public string BankName { get => bankName; set { bankName = value; OnePropertyChanged(); } }

        public string CompanyYEDROPOU { get => companyYEDROPOU; set { companyYEDROPOU = value; OnePropertyChanged(); } }

        public string AccountId { get => accountId; set { accountId = value; OnePropertyChanged(); } }

        public DateTime? AccountDate { get => accountDate; set { accountDate = value; OnePropertyChanged(); } }

        public string ActId { get => actId; set { actId = value; OnePropertyChanged(); } }

        public DateTime? ActDate { get => actDate; set { actDate = value; OnePropertyChanged(); } }

        public string CompanyDirector { get => companyDirector; set { companyDirector = value; OnePropertyChanged(); } }

        public FopModel SelectedFOP { get => FopsInfo.SelectedFop; }

        public decimal TotalAmount
        {
            get => totalAmount;
            set
            {
                totalAmount = value;
                try
                {
                    TotalAmountInWords = moneyToStrConverter.convertValue((double)value);
                }
                catch
                {
                    TotalAmountInWords = "";
                }

                OnePropertyChanged();
            }
        }

        public string TotalAmountInWords { get => totalAmountInWords; set { totalAmountInWords = value; OnePropertyChanged(); } }

        public decimal TotalAmountWithoutPDV
        {
            get => totalAmountWithoutPDV;
            set
            {
                totalAmountWithoutPDV = value;
                try
                {
                    TotalAmountWithoutPDVInWords = moneyToStrConverter.convertValue((double)value);
                }
                catch
                {
                    TotalAmountWithoutPDVInWords = "";
                }
                OnePropertyChanged();
            }
        }

        public string TotalAmountWithoutPDVInWords
        {
            get => totalAmountWithoutPDVInWords;
            set
            {
                totalAmountWithoutPDVInWords = value;
                OnePropertyChanged();
            }
        }

        public string EquipmentUsingAdress { get => equipmentUsingAdress; set { equipmentUsingAdress = value; OnePropertyChanged(); } }

        public SaveFileDialog SaveFileDialog { get; set; } = new SaveFileDialog();

        public ObservableCollection<Equipment> Equipments { get; set; }

        public List<string> GeneratedFiles { get; set; } = new List<string>();

        public ICommand GenerateDocumentCommand { get; set; }

        public ICommand ClearWindowCommand { get; set; }

        public ICommand AddEquipmentCommand { get; set; }

        public ICommand RemoveEquipmentCommand { get; set; }

        async void GenerateDocuments(object parametr = null)
        {
            try
            {
                EnableProgressBar();

                await GenerateContract();
                await GenerateAct();
                await GenerateAccount();

                OpenFiles();
            }
            catch (Exception exeption)
            {
                MessageBox.Show($"{exeption.Message}\n{exeption.StackTrace}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                DisableProgressBar();
            }

        }
        async Task GenerateContract()
        {
            var contractGenerator = new ContractGenerator();
            InitSaveFileDialog("Збереження договору", "Договір", "doc", "Word документ (.doc)|*.doc");
            if (SaveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = SaveFileDialog.FileName;
                await contractGenerator.GenerateContract(InitContractDictionary(), Equipments, filePath);
                GeneratedFiles.Add(filePath);
            }
        }
        async Task GenerateAct()
        {
            var actGenerator = new ActGenerator();
            InitSaveFileDialog("Збереження Акту", "Акт", "xlsx", "Excel документ (.xlsx)|*.xlsx");
            if (SaveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = SaveFileDialog.FileName;
                await actGenerator.GenerateAct(InitActDictionary(), filePath);
                GeneratedFiles.Add(filePath);
            }
        }

        async Task GenerateAccount()
        {
            var accountGenerator = new AccountGenerator();
            InitSaveFileDialog("Збереження рахунку", "Рахунок", "xlsx", "Excel документ (.xlsx)|*.xlsx");
            if (SaveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = SaveFileDialog.FileName;
                await accountGenerator.GenerateAccount(InitAccountDictionary(), filePath);
                GeneratedFiles.Add(filePath);
            }
        }

        Dictionary<string, string> InitContractDictionary()
        {
            return new Dictionary<string, string>()
            {
                {"ContractNumber", ContractId},
                {"ContractDate", ContractDate is null ? "-": ContractDate.Value.GetDateTimeFormats(culture)[3].Replace("р.","року")},
                {"CompanyName", CompanyName},
                {"StartDate", StartRentDate is null ? "-" : StartRentDate.Value.GetDateTimeFormats(culture)[1].Remove(5)},
                {"EndDate", EndRentDate is null ? "-" : EndRentDate.Value.ToString("dd.MM.yyyy",culture)+" року"},
                {"PostIndex", PostIndex },
                {"Adress", Adress },
                {"AccountNumber", SettlementAccount },
                {"Bank", BankName },
                {"MFO", BankMFO },
                {"CodeEDRPOY", CompanyYEDROPOU },
                {"ShortDate", ContractDate is null ? "-" : ContractDate.Value.GetDateTimeFormats(culture)[1]},
                {"FullDate", ContractDate is null ? "-": ContractDate.Value.GetDateTimeFormats(culture)[0]+"р." },
                {"Sum", TotalAmount.ToString() },
                {"EqPlace", EquipmentUsingAdress},
                {"ByWords", TotalAmountInWords},
                {"DirectoryName", CompanyDirector },
                {"FopsName", SelectedFOP.Name },
                {"Initials", SelectedFOP.Initials },
                {"FopsPostCode", SelectedFOP.PostIndex },
                {"address", SelectedFOP.Address },
                {"FopsAccountCode", SelectedFOP.AccountNumber },
                {"Fopsbank", SelectedFOP.Bank },
                {"FopsMfo", SelectedFOP.MFO },
                {"FopsIPN", SelectedFOP.CodeEDRPOY },
                {"FopsPhone", SelectedFOP.PhoneNumber  },
            };
        }

        Dictionary<string, string> InitActDictionary()
        {
            var adressPart1 = "";
            var adressPart2 = "";
            try
            {
                var words = Adress.Split(' ');
                adressPart1 = words[0] + " " + words[1];
                for (int i = 2; i < words.Length; i++)
                {
                    adressPart2 += words[i] + " ";
                }
            }
            catch
            {
                adressPart1 = "";
                adressPart2 = Adress;
            }

            return new Dictionary<string, string>()
            {
                {"ActNumber", ActId},
                {"CompanyName", CompanyName},
                {"PostIndex", PostIndex },
                {"Adress", adressPart1},
                {"AdrPart2", adressPart2},
                {"AccountNumber", SettlementAccount},
                {"Bank", BankName},
                {"MFO", BankMFO },
                {"CodeEDRPOY", CompanyYEDROPOU },
                {"ActDate", ActDate is null ? "-" : ActDate.Value.ToString("dd MMMM yyyy",culture)+" року"},
                {"FactureNumber", AccountId },
                {"FactureDate", AccountDate is null ? "-":AccountDate.Value.GetDateTimeFormats(culture)[1] },
                {"SumWithoutPDV", TotalAmountWithoutPDV.ToString("F2") },
                {"ByWords", TotalAmountWithoutPDVInWords},
                {"fopsinitials", SelectedFOP.Initials },
                {"fopspostindex", SelectedFOP.PostIndex },
                {"fopsaddress", SelectedFOP.Address },
                {"fopsaccountnumber", SelectedFOP.AccountNumber },
                {"fopsbank", SelectedFOP.Bank },
                {"fopsmfo", SelectedFOP.MFO },
                {"fopscodeedrpoy", SelectedFOP.CodeEDRPOY },
            };
        }

        Dictionary<string, string> InitAccountDictionary()
        {
            int days = 0;
            if (StartRentDate != null && EndRentDate != null)
            {
                days = (EndRentDate - StartRentDate).Value.Days;
            }
            return new Dictionary<string, string>()
            {
                {"CompanyName", CompanyName},
                {"PostIndex", PostIndex },
                {"Adress", Adress },
                {"AccountNumber", SettlementAccount },
                {"Bank", BankName },
                {"MFO", BankMFO },
                {"CodeEDRPOY", CompanyYEDROPOU },
                {"FactureNumber", AccountId },
                {"FactureDate", AccountDate is null ? "-" : AccountDate.Value.GetDateTimeFormats(culture)[1] },
                {"StartRentDate", $"{StartRentDate.Value.ToString("dd.MM")} - {EndRentDate.Value.ToString("dd.MM")}" },
                {"SumWithoutPDV", TotalAmountWithoutPDV.ToString("F2") },
                {"ByWords", TotalAmountWithoutPDVInWords},
                {"CountDay", days.ToString() },
                {"fopsname", SelectedFOP.Name },
                {"fopspostindex", SelectedFOP.PostIndex },
                {"fopsaddress", SelectedFOP.Address },
                {"fopsaccountnumber", SelectedFOP.AccountNumber },
                {"fopsbank", SelectedFOP.Bank },
                {"fopsmfo", SelectedFOP.MFO },
                {"fopscodeedrpoy", SelectedFOP.CodeEDRPOY },
            };
        }

        void InitSaveFileDialog(string title, string fileName, string defExtension, string filter)
        {
            SaveFileDialog.Title = title;
            SaveFileDialog.FileName = fileName;
            SaveFileDialog.DefaultExt = defExtension;
            SaveFileDialog.Filter = filter;
        }
        void ClearWindow(object parametr = null)
        {
            ContractId = "";
            AccountDate = null;
            AccountId = "";
            ActDate = null;
            ActId = "";
            Adress = "";
            BankMFO = "";
            BankName = "";
            CompanyDirector = "";
            CompanyName = "";
            CompanyYEDROPOU = "";
            ContractDate = null;
            StartRentDate = null;
            EndRentDate = null;
            TotalAmount = 0.0M;
            TotalAmountInWords = "";
            TotalAmountWithoutPDV = 0.0M;
            TotalAmountWithoutPDVInWords = "";
            PostIndex = "";
            EquipmentUsingAdress = "";
            SettlementAccount = "";
            Equipments.Clear();
        }
        void AddEquipment(object parametr = null)
        {
            try
            {
                var timeSpan = EndRentDate - StartRentDate;
                Equipments.Add(new Equipment()
                {
                    Termin = timeSpan.Value.Days,
                    StartDate = StartRentDate.Value.ToString("dd/MM/yy"),
                    EndDate = EndRentDate.Value.ToString("dd/MM/yy")
                });
            }
            catch (Exception)
            {
                Equipments.Add(new Equipment());
            }
        }

        void AddDefultEquipment()
        {
            string strartDate;
            string endDate;
            int termin = 0;
            DocumentData documentData = null;
            try
            {
                BinaryFormatter formatter = new BinaryFormatter();
                using (FileStream fs = new FileStream("data.dat", FileMode.Open, FileAccess.Read))
                {
                    documentData = (DocumentData)formatter.Deserialize(fs);
                }
            }
            catch (Exception)
            {
                //ignore
            }
           
            strartDate = documentData?.StartRentDate?.ToString("dd/MM/yy");
            endDate = documentData?.EndRentDate?.ToString("dd/MM/yy");
            if (documentData?.EndRentDate != null && documentData?.StartRentDate != null)
            {
                var timeSpan = documentData.EndRentDate - documentData.StartRentDate;
                termin = timeSpan.Value.Days;
            }

            Equipments.Add(new Equipment()
            {
                Name = "Комплекс послуг з оренди меблів та обладнання",
                Termin = termin,
                StartDate = strartDate,
                EndDate = endDate
            });

            for (int i = 0; i < 2; i++)
            {
                Equipments.Add(new Equipment());
            }
        }
        void RemoveEquipment(object parametr = null)
        {
            int selectedIndex = (int)parametr;
            if (selectedIndex != -1)
            {
                Equipments.RemoveAt(selectedIndex);
            }
        }

        void OpenFiles()
        {
            foreach (var file in GeneratedFiles)
            {
                try
                {
                    Process.Start(file);
                }
                catch
                {
                    //ignore
                }
            }
            GeneratedFiles.Clear();
        }
    }
}
