using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class Labour
    {
        public int? Id { get; set; }
        public string Groupe { get; set; }
        public string Item { get; set; }
        public string Description { get; set; }
        public string Contractor { get; set; }
        public decimal Quantity { get; set; }
        public decimal Rate { get; set; }
        public decimal Adjust { get; set; }
        public decimal Price { get; set; }
        public int Percent { get; set; }
        public decimal Payout { get; set; }
        public decimal Profit { get; set; }        
        
        public int? LabourProfitId { get; set; }
        public virtual LabourProfit LabourProfit  { get; set; }
    }
}
