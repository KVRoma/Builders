using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class Invoice
    {
        public int? Id { get; set; }
        public string NumberInvoice 
        {
            get { return "I" + DateInvoice.Year.ToString() + 
                               ((DateInvoice.Month < 10) ? ("0" + DateInvoice.Month.ToString()) : (DateInvoice.Month.ToString())) + 
                               ((DateInvoice.Day < 10) ? ("0" + DateInvoice.Day.ToString()) : (DateInvoice.Day.ToString())) + 
                               NumberQuota; }              
        }
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
            get { return NumberQuota + " " + FirstName + " " + LastName + " " + PhoneNumber + " " + Email; }
        }
        public string NumberComboBox
        {
            get { return NumberInvoice + " " + FullName; }
        }
        public DateTime DateInvoice { get; set; }
        public int? QuotaId { get; set; }
        public string NumberQuota { get; set; }        
        public string UpNumber { get; set; }
        public string OrderNumber { get; set; }

        public decimal MaterialProfit { get; set; }
        public decimal LabourProfit { get; set; }
        public decimal TotalProfit { get; set; }

        public string CompanyName { get; set; } = "CMO";
    }
}
