using SQLite.CodeFirst;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class Quotation
    {        
        public int? Id { get; set; }
        public string PrefixNumberQuota { get; set; } // Q and IP and RQ
        public string NumberQuota 
        {
            get { return PrefixNumberQuota + (Id + 1000).ToString(); } 
        } 
        public string NumberComboBox
        {
            get { return NumberQuota + " " + FullName; }
        }
        public string NumberClient { get; set; }
        public DateTime QuotaDate { get; set; }
        public int ClientId { get; set; }

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
            get { return NumberClient + " " + FirstName + " " + LastName + " " + PhoneNumber + " " + Email; }
        }
        public string JobDescription { get; set; }
        public string JobNote { get; set; }

        public decimal MaterialSubtotal { get; set; }
        public decimal MaterialTax 
        {
            get { return decimal.Round(MaterialSubtotal * 0.12m, 2); }            
        }        
        public decimal MaterialDiscountYN { get; set; }
        public decimal MaterialDiscountAmount { get; set; }
        public decimal MaterialTotal 
        {
            get { return (MaterialSubtotal + MaterialTax) - MaterialDiscountAmount; }
        }

        public decimal LabourSubtotal { get; set; }
        public decimal LabourTax 
        {
            get {return decimal.Round(LabourSubtotal * 0.05m, 2); } 
        }
        public decimal LabourDiscountYN { get; set; }
        public decimal LabourDiscountAmount { get; set; }
        public decimal LabourTotal 
        {
            get { return (LabourSubtotal + LabourTax) - LabourDiscountAmount; }
        }

        public decimal ProjectTotal 
        {
            get{return MaterialTotal + LabourTotal; }
        }
        public bool FinancingYesNo { get; set; }
        public decimal FinancingAmount 
        {
            get { return (FinancingYesNo) ? (ProjectTotal) : (0m); }
        }
        public decimal FinancingFee 
        {
            get { return decimal.Round((FinancingYesNo) ? (FinancingAmount * 0.08m) : (0m),2); }
        }
        public decimal AmountPaidByCreditCard { get; set; }
        public decimal ProcessingFee 
        {
            get { return decimal.Round(AmountPaidByCreditCard * 0.03m, 2); }
        }
        public decimal InvoiceGrandTotal 
        {
            get { return ProjectTotal + FinancingFee + ProcessingFee; }
        }

        public bool ActivQuota { get; set; } = false;
        public bool PaidQuota { get; set; } = false;
        public int SortingQuota { get; set; } = 3; // creation - 3, active - 1, paid - 2
        public string Color { get; set; } = "Black";
       

        public ICollection<MaterialQuotation> MaterialQuotation { get; set; }
        public ICollection<Payment> Payment { get; set; }
        public Quotation()
        {            
            MaterialQuotation = new List<MaterialQuotation>();
            Payment = new List<Payment>();
        }
    }
}
