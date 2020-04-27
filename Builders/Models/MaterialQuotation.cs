using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class MaterialQuotation
    {
        public int? Id { get; set; }
        public string Groupe { get; set; }
        public string Item { get; set; }
        public string Description { get; set; }
        public decimal Quantity { get; set; }
        public decimal Rate { get; set; }
        public decimal Price { get; set; }        

        public int? QuotationId { get; set; }
        public Quotation Quotation { get; set; }
    }
}
