using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class ReportTotal
    {        
        public decimal TotalSale { get; set; }
        public decimal TotalMaterialAndLabourCost { get; set; }
        public decimal TotalCrossProfit 
        { 
            get { return decimal.Round(TotalSale - TotalMaterialAndLabourCost, 2); }
        }
        public decimal TotalExpenses { get; set; }
        public decimal TotalProcessFees { get; set; }
        public decimal NetStoreProfit 
        {
            get { return decimal.Round(TotalCrossProfit - TotalExpenses + TotalProcessFees, 2); } 
        }
        public decimal NetStoreProfitHalf 
        { 
            get { return decimal.Round(NetStoreProfit / 2m, 2); } 
        }
        public decimal NetProfit { get; set; }
        public decimal MaterialSplit 
        { 
            get { return decimal.Round(NetProfit / 2m, 2); } 
        }
    }
}
