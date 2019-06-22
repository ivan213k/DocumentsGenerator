using System;
using System.Collections.ObjectModel;
using System.Windows.Input;
using DocumentsGenerator.Models;

namespace DocumentsGenerator.ViewModels
{
    class MainViewModel: BaseViewModel
    {
        public MainViewModel()
        {
            Equipments = new ObservableCollection<Equipment>();
            GenerateDocumentCommand = new Command(GenerateDocuments);
            ClearWindowCommand = new Command(ClearWindow);
            AddEquipmentCommand = new Command(AddEquipment);
            RemoveEquipmentCommand = new Command(RemoveEquipment);
        }

        public string ContractId { get; set; }

        public DateTime? ContractDate { get; set; } 

        public string CompanyName { get; set; }

        public DateTime? StartRentDate { get; set; }

        public DateTime? EndRentDate { get; set; }

        public string PostIndex { get; set; }

        public string Adress { get; set; }

        public string SettlementAccount { get; set; }

        public string BankMFO { get; set; }

        public string BankName { get; set; }

        public string CompanyYEDROPOU { get; set; }

        public string AccountId { get; set; }

        public DateTime? AccountDate { get; set; }

        public string ActId { get; set; }

        public DateTime? ActDate { get; set; }

        public string CompanyDirector { get; set; }

        public decimal? TotalAmount { get; set; }

        public string TotalAmountInWords { get; set; }

        public decimal? TotalAmountWithoutPDV { get; set; }

        public string TotalAmountWithoutPDVInWords { get; set; }

        public string EquipmentUsingAdress { get; set; }

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
            if (selectedIndex!=-1)
            {
                Equipments.RemoveAt(selectedIndex);
            }
        }
    }
}
