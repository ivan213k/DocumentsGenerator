using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Input;
using DocumentsGenerator.Models;

namespace DocumentsGenerator.ViewModels
{
    class MainViewModel : BaseViewModel
    {
        private string contractId;
        private DateTime? contractDate;
        private string companyName;
        private DateTime? startRentDate;
        private DateTime? endRentDate;
        private string postIndex;
        private string adress;
        private string settlementAccount;
        private string bankMFO;
        private string bankName;
        private string companyYEDROPOU;
        private string accountId;
        private DateTime? accountDate;
        private string actId;
        private DateTime? actDate;
        private string companyDirector;
        private decimal totalAmount;
        private string totalAmountInWords;
        private decimal totalAmountWithoutPDV;
        private string totalAmountWithoutPDVInWords;
        private string equipmentUsingAdress;
        MoneyToStrConverter moneyToStrConverter = new MoneyToStrConverter("UAH", "UKR", "F");
        public MainViewModel()
        {
            Equipments = new ObservableCollection<Equipment>();
            GenerateDocumentCommand = new Command(GenerateDocuments);
            ClearWindowCommand = new Command(ClearWindow);
            AddEquipmentCommand = new Command(AddEquipment);
            RemoveEquipmentCommand = new Command(RemoveEquipment);
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
            GeneratedFiles.Clear();
            EnableProgressBar();

            await GenerateContract();
            await GenerateAct();
            await GenerateAccount();

            DisableProgressBar();
            OpenFiles(GeneratedFiles);
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
                {"ContractNumber",ContractId},
                {"ContractDate",ContractDate == null ? "-": ContractDate.Value.GetDateTimeFormats()[3].Replace("р.","року")},
                {"CompanyName",CompanyName},
                {"StartDate",StartRentDate==null ? "-" : StartRentDate.Value.GetDateTimeFormats()[1].Remove(5)},
                {"EndDate",EndRentDate==null ? "-" : EndRentDate.Value.ToString("dd.MM.yyyy")+" року"},
                {"PostIndex",PostIndex },
                {"Adress",Adress },
                {"AccountNumber",SettlementAccount },
                {"Bank",BankName },
                {"MFO",BankMFO },
                {"CodeEDRPOY",CompanyYEDROPOU },
                {"ShortDate",ContractDate==null? "-" : ContractDate.Value.GetDateTimeFormats()[1]},
                {"FullDate",ContractDate ==null? "-": ContractDate.Value.GetDateTimeFormats()[0]+"р." },
                {"Sum",TotalAmount.ToString() },
                {"EqPlace",EquipmentUsingAdress},
                {"ByWords",TotalAmountInWords},
                {"DirectoryName",CompanyDirector }
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
                {"ActNumber",ActId },
                {"CompanyName",CompanyName},
                {"PostIndex",PostIndex },
                {"Adress",adressPart1 },
                {"AdrPart2",adressPart2 },
                {"AccountNumber",SettlementAccount },
                {"Bank",BankName },
                {"MFO",BankMFO },
                {"CodeEDRPOY",CompanyYEDROPOU },
                {"ActDate" ,ActDate==null?"-":ActDate.Value.ToString("dd MMMM yyyy")+" року"},
                {"FactureNumber",AccountId },
                {"FactureDate",AccountDate==null?"-":AccountDate.Value.GetDateTimeFormats()[1] },
                {"SumWithoutPDV",TotalAmountWithoutPDV.ToString("F2") },
                {"ByWords",TotalAmountWithoutPDVInWords},
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
                {"CompanyName",CompanyName},
                {"PostIndex",PostIndex },
                {"Adress",Adress },
                {"AccountNumber",SettlementAccount },
                {"Bank",BankName },
                {"MFO",BankMFO },
                {"CodeEDRPOY",CompanyYEDROPOU },
                {"FactureNumber",AccountId },
                {"FactureDate",AccountDate==null ? "-" : AccountDate.Value.GetDateTimeFormats()[1] },
                {"SumWithoutPDV",TotalAmountWithoutPDV.ToString("F2") },
                {"ByWords",TotalAmountWithoutPDVInWords},
                {"CountDay", days.ToString() }
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
        void RemoveEquipment(object parametr = null)
        {
            int selectedIndex = (int)parametr;
            if (selectedIndex != -1)
            {
                Equipments.RemoveAt(selectedIndex);
            }
        }

        void OpenFiles(IEnumerable<string> files)
        {
            foreach (var file in files)
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
        }
    }
}
