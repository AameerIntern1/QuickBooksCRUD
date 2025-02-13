using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuickBooksCRUD
{
    public class JournalEntry
    {
       public string? Account { get; set; }
        public double? EarnedAmount { get; set; }
        public double? UnEarnedAmount { get; set; }
        public double? AccountReceivable { get; set; }
        public double? Cash { get; set; }
    }
}
