using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class WorkOrder_Installation
    {
        public int? Id { get; set; }
        public string Groupe { get; set; }
        public string Item { get; set; }
        public string Description { get; set; }
        public decimal Quantity { get; set; }
        public decimal Rate { get; set; }
        public decimal Price { get; set; }
        public decimal Procent { get; set; }
        public decimal Payout { get; set; }
        public string Contractor { get; set; }
       
        public string Color { get; set; }

        public int? WorkOrderId { get; set; }
        public virtual WorkOrder WorkOrder { get; set; }
    }
}
