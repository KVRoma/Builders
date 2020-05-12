using Builders.Commands;
using Builders.Models;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.ViewModels
{
    public class MaterialProfitViewModel : ViewModel
    {
        public string WindowName { get; } = "Material Profit";
        private BuilderContext db;
        
        private MaterialProfit profit;
        private Material materialSelect;
        private IEnumerable<Material> materials;
        private Material expSelect;
        private IEnumerable<Material> exps;

        private decimal quantity;
        private decimal unitPrice;
        private decimal epRate;
        private decimal tax;

        private string description;
        private decimal otherQuantity;
        private decimal otherUnitPrice;
        private decimal otherEpRate;
        private decimal otherTax;

        public MaterialProfit Profit
        {
            get { return profit; }
            set
            {
                profit = value;
                OnPropertyChanged(nameof(Profit));
            }
        }
        public Material MaterialSelect
        {
            get { return materialSelect; }
            set
            {
                materialSelect = value;
                OnPropertyChanged(nameof(MaterialSelect));                
            }
        }
        public IEnumerable<Material> Materials
        {
            get { return materials; }
            set
            {
                materials = value;
                OnPropertyChanged(nameof(Materials));
            }
        }
        public Material ExpSelect
        {
            get { return expSelect; }
            set
            {
                expSelect = value;
                OnPropertyChanged(nameof(ExpSelect));
                if (ExpSelect != null)
                {
                    Description = ExpSelect.Description;
                    OtherQuantity = ExpSelect.CostQuantity;
                    OtherUnitPrice = ExpSelect.CostUnitPrice;
                    OtherEpRate = ExpSelect.CostEPRate;
                    OtherTax = ExpSelect.CostTax;
                }
            }
        }
        public IEnumerable<Material> Exps
        {
            get { return exps; }
            set
            {
                exps = value;
                OnPropertyChanged(nameof(Exps));
            }
        }

        public decimal Quantity
        {
            get { return quantity; }
            set
            {
                quantity = decimal.Round(value, 2);
                OnPropertyChanged(nameof(Quantity));
            }
        }
        public decimal UnitPrice
        {
            get { return unitPrice; }
            set
            {
                unitPrice = decimal.Round(value, 2);
                OnPropertyChanged(nameof(UnitPrice));
            }
        }
        public decimal EpRate
        {
            get { return epRate; }
            set
            {
                epRate = decimal.Round(value, 2);
                OnPropertyChanged(nameof(EpRate));
            }
        }
        public decimal Tax
        {
            get { return tax; }
            set
            {
                tax = decimal.Round(value, 2);
                OnPropertyChanged(nameof(Tax));
            }
        }

        public string Description
        {
            get { return description; }
            set            
            {
                description = value;
                OnPropertyChanged(nameof(Description));
            }
        }
        public decimal OtherQuantity
        {
            get { return otherQuantity; }
            set
            {
                otherQuantity = decimal.Round(value, 2);
                OnPropertyChanged(nameof(OtherQuantity));
            }
        }
        public decimal OtherUnitPrice
        {
            get { return otherUnitPrice; }
            set
            {
                otherUnitPrice = decimal.Round(value, 2);
                OnPropertyChanged(nameof(OtherUnitPrice));
            }
        }
        public decimal OtherEpRate
        {
            get { return otherEpRate; }
            set
            {
                otherEpRate = decimal.Round(value, 2);
                OnPropertyChanged(nameof(OtherEpRate));
            }
        }
        public decimal OtherTax
        {
            get { return otherTax; }
            set
            {
                otherTax = decimal.Round(value, 2);
                OnPropertyChanged(nameof(OtherTax));
            }
        }
        //************************************************************
        private Command _insMaterial;
        private Command _addExps;
        private Command _insExps; 
        private Command _delExps;

        public Command InsMaterial => _insMaterial ?? (_insMaterial = new Command(obj=> 
        {
            if (MaterialSelect != null)
            {                
                decimal subTotal = decimal.Round(Quantity * UnitPrice, 2);
                MaterialSelect.CostQuantity = Quantity;
                MaterialSelect.CostUnitPrice = decimal.Round(UnitPrice, 2);
                MaterialSelect.CostEPRate = EpRate;
                MaterialSelect.CostTax = Tax;
                MaterialSelect.CostSubtotal = subTotal;
                MaterialSelect.CostTotal = decimal.Round(subTotal + (subTotal * (Tax / 100m)) - (subTotal * EpRate), 2);
                db.Entry(MaterialSelect).State = EntityState.Modified;
                db.SaveChanges();
                Materials = null;
                Materials = db.Materials.Local.ToBindingList().Where(m => m.MaterialProfitId == profit.Id).Where(g => g.Groupe != "OTHER");                
            }
        }));
        public Command AddExps => _addExps ?? (_addExps = new Command(obj=> 
        {
            if (Description != "")
            {
                decimal CostSubTotal = decimal.Round(OtherQuantity * OtherUnitPrice, 2);
                Material mat = new Material() 
                {
                    MaterialProfitId = profit.Id,
                    Groupe = "OTHER",
                    Item = "Item #" + (Exps.Count() + 1).ToString(),
                    Description = Description,
                    Quantity = 0m,
                    Rate = 0m,
                    Price = 0m,
                    CostQuantity = OtherQuantity,
                    CostUnitPrice = OtherUnitPrice,
                    CostEPRate = OtherEpRate,
                    CostTax = OtherTax,
                    CostSubtotal = CostSubTotal,
                    CostTotal = decimal.Round(CostSubTotal + (CostSubTotal * (OtherTax/100m)) - (CostSubTotal * OtherEpRate), 2)
                };
                db.Materials.Add(mat);
                db.SaveChanges();
                Exps = null;
                Exps = db.Materials.Local.ToBindingList().Where(m => m.MaterialProfitId == profit.Id).Where(g => g.Groupe == "OTHER");
            }
        }));
        public Command InsExps => _insExps ?? (_insExps = new Command(obj=> 
        {
            if (ExpSelect != null)
            {
                decimal CostSubTotal = decimal.Round(OtherQuantity * OtherUnitPrice, 2);
                ExpSelect.CostQuantity = OtherQuantity;
                ExpSelect.CostUnitPrice = OtherUnitPrice;
                ExpSelect.CostEPRate = OtherEpRate;
                ExpSelect.CostTax = OtherTax;
                ExpSelect.CostSubtotal = CostSubTotal;
                ExpSelect.CostTotal = decimal.Round(CostSubTotal + (CostSubTotal * (OtherTax/100m)) - (CostSubTotal * OtherEpRate), 2);

                db.Entry(ExpSelect).State = EntityState.Modified;
                db.SaveChanges();
                Exps = null;
                Exps = db.Materials.Local.ToBindingList().Where(m => m.MaterialProfitId == profit.Id).Where(g => g.Groupe == "OTHER");
            }
        }));
        public Command DelExps => _delExps ?? (_delExps = new Command(obj=> 
        {
            if (ExpSelect != null)
            {
                db.Materials.Remove(ExpSelect);
                db.SaveChanges();
                Exps = null;
                Exps = db.Materials.Local.ToBindingList().Where(m => m.MaterialProfitId == profit.Id).Where(g => g.Groupe == "OTHER");
            }
        }));
        //************************************************************

        public MaterialProfitViewModel(ref BuilderContext context, MaterialProfit select)
        {
            db = context;
            profit = select;
            db.Materials.Load();
            Materials = db.Materials.Local.ToBindingList().Where(m => m.MaterialProfitId == profit.Id).Where(g=>g.Groupe != "OTHER");
            Exps = db.Materials.Local.ToBindingList().Where(m => m.MaterialProfitId == profit.Id).Where(g => g.Groupe == "OTHER");
        }
    }
}
