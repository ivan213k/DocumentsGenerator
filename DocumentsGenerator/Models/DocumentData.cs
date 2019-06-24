using System;

namespace DocumentsGenerator.Models
{
    [Serializable]
    class DocumentData
    {
        public string ContractId { get; set; }

        public DateTime? ContractDate { get; set; }

        public string CompanyName { get; set; }

        public DateTime? StartRentDate { get; set; }

        public DateTime? EndRentDate { get; set; }

        public string PostIndex { get; set; }

        public string Adress { get; set; }

        public string AccountId { get; set; }

        public string MFO { get; set; }

        public decimal? AmountWithoutPDV { get; set; }

        public string AmountWithoutPDVInWords { get; set; }

        public decimal? TotalAmount { get; set; }

        public string TotalAmountInWords { get; set; }

        public string EDROPOY { get; set; }

        public string FacktureId { get; set; }

        public DateTime? AccountDate { get; set; }

        public string ActId { get; set; }

        public DateTime? ActDate { get; set; }

        public string Bank { get; set; }

        public string EquipmentUsingAdress { get; set; }

        public string DirectoryName { get; set; }
    }
}
