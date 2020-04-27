using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class LabourProfit
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

        public decimal CollectedSubtotal { get; set; }
        public decimal CollectedGST { get; set; }
        public decimal CollectedTotal { get; set; }
        
        public decimal PayoutSubtotal { get; set; }
        public decimal PayoutGST { get; set; }
        public decimal PayoutTotal { get; set;}
        
        public decimal StoreSubtotal { get; set; }
        public decimal StoreGST { get; set; }
        public decimal StoreTotal { get; set; }

        public decimal Discount { get; set; }
        public decimal ProfitTotal { get; set; }

        public virtual ICollection<Labour> Labours  { get; set; }
        public virtual ICollection<LabourContractor> LabourContractors { get; set; }
        public LabourProfit()
        {
            Labours = new List<Labour>();
            LabourContractors = new List<LabourContractor>();
        }
    }
}
