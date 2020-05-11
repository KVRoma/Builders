using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class Payment
    {
        public int? Id { get; set; }
        public int? NumberPayment { get; set; } 
        public DateTime PaymentDatePaid { get; set; }
        public decimal PaymentAmountPaid { get; set; }
        public decimal PaymentPrincipalPaid { get; set; }
        public string PaymentMethod { get; set; }
        public decimal ProcessingFee { get; set; }
        public decimal Balance { get; set; }

        public int? QuotationId { get; set; }
        public Quotation Quotation { get; set; }
    }
}
