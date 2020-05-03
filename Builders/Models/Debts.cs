using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class Debts
    {
        public int Id { get; set; }
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
        public string NumberComboBox
        {
            get { return InvoiceNumber + " " + FullName; }
        }

        public string NameDebts { get; set; }
        public string DescriptionDebts { get; set; }
        public decimal AmountDebts { get; set; }

        public DateTime DatePayment { get; set; }
        public decimal AmountPayment { get; set; }
        public string DescriptionPayment { get; set; }

        public bool Payment { get; set; } = false;
        public string ColorPayment { get; set; } = "Red";
        public bool ReadOnly { get; set; } = true;
    }
}
