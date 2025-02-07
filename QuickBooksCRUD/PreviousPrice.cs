using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuickBooksCRUD
{
    public class PreviousPrice
    {
        public string? Id { get; set; }
        public string? TaxId { get; set; }
        public string? EditSequenceID { get; set; }
        public string? Item { get; set; }
        public decimal? OldPrice { get; set; }
        public decimal? NewPrice { get; set; }
        public DateTime? TxnDate { get; set; }
    }

}