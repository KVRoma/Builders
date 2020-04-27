using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class ReportPayrollToWork
    {
        //public int? WorkOrderId { get; set; }
        //public int? QuotaId { get; set; }
        public string NumberInvoice { get; set; }
        
        public DateTime DateInvoice { get; set; }
        //public DateTime DateServices { get; set; }
        //public DateTime DateCompletion { get; set; }
        public string Contractor { get; set; }
        public decimal Payout { get; set; }
        public decimal Adjust { get; set; }
        public decimal Total { get; set; }
        public string TAX { get; set; }
        public decimal GST { get; set; }
        public decimal TotalContractor { get; set; }
    }
}
