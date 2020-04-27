using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class WorkOrder_Contractor
    {
        public int? Id { get; set; }
        public string Contractor { get; set; }
        public decimal Payout { get; set; }
        public decimal Adjust { get; set; }
        public decimal Total { get; set; }
        public string TAX { get; set; }
        public decimal GST { get; set; }
        public decimal TotalContractor { get; set; }
        
        public string Color { get; set; }

        public int? WorkOrderId { get; set; }
        public virtual WorkOrder WorkOrder { get; set; }
    }
}
