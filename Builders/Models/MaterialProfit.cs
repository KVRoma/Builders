using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class MaterialProfit
    {
        public int? Id { get; set; }
        public int? InvoiceId { get; set; }
        public DateTime InvoiceDate { get; set; }
        public string InvoiceNumber { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string PhoneNumber { get; set; }
        public string Email { get; set; }
        public string FullName
        {
            get { return FirstName + " " + LastName; }
        }
        public string FullSearch
        {
            get { return InvoiceNumber + " " + FirstName + " " + LastName + " " + PhoneNumber + " " + Email; }
        }

        public decimal MaterialSubtotal { get; set; }
        public decimal MaterialTax { get; set; }
        public decimal MaterialTotal { get; set; }
        
        public decimal CostMaterialSubtotal { get; set; }
        public decimal CostMaterialTax { get; set; }
        public decimal CostMaterialTotal { get; set; }
        
        public decimal ProfitBeforTax { get; set; }
        public decimal ProfitTax { get; set; }
        public decimal ProfitInclTax { get; set; }
        public decimal ProfitDiscount { get; set; }
        public decimal ProfitTotal { get; set; }

        public bool Companion { get; set; } = true;

        public string CompanyName { get; set; } = "CMO";

        public virtual ICollection<Material> Materials { get; set; }
        public MaterialProfit()
        {
            Materials = new List<Material>();           
        }
    }
}
