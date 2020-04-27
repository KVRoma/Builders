using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class Material
    {
        public int? Id { get; set; }
        public string Groupe { get; set; }
        public string Item { get; set; }
        public string Description { get; set; }
        public decimal Quantity { get; set; }
        public decimal Rate { get; set; }
        public decimal Price { get; set; }
        public decimal CostQuantity { get; set; }
        public decimal CostUnitPrice { get; set; }
        public decimal CostSubtotal { get; set; }
        public decimal CostEPRate { get; set; }
        public decimal CostTax { get; set; }
        public decimal CostTotal { get; set; }       

        public int? MaterialProfitId { get; set; }
        public virtual MaterialProfit MaterialProfit { get; set; }
    }
}
