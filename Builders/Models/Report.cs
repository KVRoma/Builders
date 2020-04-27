using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class Report
    {
       
        public int Id { get; set; }        
        public string InvoiceNumber { get; set; }
        public DateTime InvoiceDate { get; set; }
        public string PrimaryName { get; set; }
        public decimal MaterialSubtotal { get; set; }
        public decimal MaterialTax { get; set; }
        public decimal MaterialGrandTotal { get; set; }

        public decimal LabourSubtotal { get; set; }
        public decimal LabourGST { get; set; }
        public decimal LabourGrandTotal { get; set; }

        public decimal MaterialAndLabourGrandTotal { get; set; }

        public decimal MaterialCostSubtotal { get; set; }
        public decimal MaterialCostTax { get; set; }
        public decimal MaterialDiscountDeductions { get; set; }
        public decimal MaterialCostTotal { get; set; }

        public decimal LabourPayoutTotalBeforeGST { get; set; }
        public decimal LabourPayoutGST { get; set; }
        public decimal LabourDiscountDeductions { get; set; }
        public decimal LabourPayoutTotalAfterGST { get; set; }

        public decimal TotalMaterialProfits { get; set; }
        public decimal TotalLabourProfits { get; set; }
        public decimal ProcessingFeeCollected { get; set; }

        public decimal TotalProfit { get; set; }
        
    }
}
