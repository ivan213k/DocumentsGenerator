using System;
using System.Collections.ObjectModel;
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
        private decimal? totalAmount;
        private string totalAmountInWords;
        private decimal? totalAmountWithoutPDV;
        private string totalAmountWithoutPDVInWords;
        private string equipmentUsingAdress;

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

        public decimal? TotalAmount { get => totalAmount; set { totalAmount = value; OnePropertyChanged(); } }

        public string TotalAmountInWords { get => totalAmountInWords; set { totalAmountInWords = value; OnePropertyChanged(); } }

        public decimal? TotalAmountWithoutPDV { get => totalAmountWithoutPDV; set { totalAmountWithoutPDV = value; OnePropertyChanged(); } }

        public string TotalAmountWithoutPDVInWords { get => totalAmountWithoutPDVInWords; set { totalAmountWithoutPDVInWords = value; OnePropertyChanged(); } }

        public string EquipmentUsingAdress { get => equipmentUsingAdress; set { equipmentUsingAdress = value; OnePropertyChanged(); } }

        public ObservableCollection<Equipment> Equipments { get; set; }


        public ICommand GenerateDocumentCommand { get; set; }

        public ICommand ClearWindowCommand { get; set; }

        public ICommand AddEquipmentCommand { get; set; }

        public ICommand RemoveEquipmentCommand { get; set; }

        void GenerateDocuments(object parametr = null)
        {

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
            TotalAmount = null;
            TotalAmountInWords = "";
            TotalAmountWithoutPDV = null;
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
    }
}
