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

        public Dictionary<int, string> GetValues()
        {
            Dictionary<int, string> cellIndexPairs = new Dictionary<int, string>();
            cellIndexPairs.Add(1, Name);
            cellIndexPairs.Add(2, Count.ToString());
            cellIndexPairs.Add(3, ReplacmentCost.ToString());
            cellIndexPairs.Add(4, Termin.ToString());
            cellIndexPairs.Add(5, StartDate);
            cellIndexPairs.Add(6, EndDate);
            cellIndexPairs.Add(7, Amount.ToString());
            return cellIndexPairs;
        }
    }
}
