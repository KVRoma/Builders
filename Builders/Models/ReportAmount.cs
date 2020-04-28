using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class ReportAmount
    {
        public DateTime QuotaDate { get; set; }
        public string NumberQuota { get; set; }
        public string FullNameQuota { get; set; }
        public DateTime PaymentDatePaid { get; set; }
        public decimal PaymentAmountPaid { get; set; }
        public string PaymentMethod { get; set; }
        public decimal Balance { get; set; }
    }
}
