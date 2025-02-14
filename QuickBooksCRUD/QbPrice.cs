using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuickBooksCRUD
{
    public class QbPrice
    {
        public string? Id { get; set; }
        public string? TaxId { get; set; }
        public string? CreditTxnLineId { get; set; }
        public string? DebitTxnLineId { get; set; }
        public string? EditSequenceID { get; set; }
        public string? CreditAccount { get; set; }
        public string? DebitAccount { get; set; }
        public decimal? DebitPrice { get; set; }
        public decimal? CreditPrice { get; set; }
        public DateTime? TxnDate { get; set; }
    }
}






