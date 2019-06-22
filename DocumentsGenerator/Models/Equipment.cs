using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentsGenerator.Models
{
    class Equipment
    {
        public string Name { get; set; }

        public int Count { get; set; }

        public decimal ReplacmentCost { get; set; }

        public int Termin { get; set; }

        public string StartDate { get; set; }

        public string EndDate { get; set; }

        public decimal Amount { get; set; }
    }
}
