using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class ReportPayroll
    {
        public string Contractor { get; set; }
        public decimal Payout { get; set; }
        public decimal Adjust { get; set; }
        public decimal Total { get; set; }        
        public decimal GST { get; set; }
        public decimal TotalContractor { get; set; }
    }
}
