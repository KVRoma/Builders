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
    public class LabourProfitViewModel : ViewModel
    {
        public string WindowName { get; } = "Labour Profit";
        private BuilderContext db;
        private LabourProfit labourProfit;
        private Labour labourSelect;
        private IEnumerable<Labour> labours;
        private LabourContractor labourContractorSelect;
        private IEnumerable<LabourContractor> labourContractors;
        private DIC_Contractor contractorSelect;
        private IEnumerable<DIC_Contractor> contractors;
        private decimal percent;
        private decimal adjust;
        private string taxSelect;
        private List<string> tax;
        private decimal adjustContractor;

        public LabourProfit LabourProfit
        {
            get { return labourProfit; }
            set
            {
                labourProfit = value;
                OnPropertyChanged(nameof(LabourProfit));
            }
        }
        public Labour LabourSelect
        {
            get { return labourSelect; }
            set
            {
                labourSelect = value;
                OnPropertyChanged(nameof(LabourSelect));
                if (LabourSelect != null)
                {
                    ContractorSelect = db.DIC_Contractors.FirstOrDefault(c => c.Name == LabourSelect.Contractor);
                    Percent = LabourSelect.Percent;
                    Adjust = LabourSelect.Adjust;
                }
            }
        }
        public IEnumerable<Labour> Labours
        {
            get { return labours; }
            set 
            {
                labours = value;
                OnPropertyChanged(nameof(Labours));
            }
        }
        public LabourContractor LabourContractorSelect
        {
            get { return labourContractorSelect; }
            set
            {
                labourContractorSelect = value;
                OnPropertyChanged(nameof(LabourContractorSelect));
                if (LabourContractorSelect != null)
                {
                    AdjustContractor = LabourContractorSelect.Adjust;
                    TaxSelect = LabourContractorSelect.TAX;
                }
            }
        }
        public IEnumerable<LabourContractor> LabourContractors
        {
            get { return labourContractors; }
            set
            {
                labourContractors = value;
                OnPropertyChanged(nameof(LabourContractors));
            }
        }
        public DIC_Contractor ContractorSelect
        {
            get { return contractorSelect; }
            set
            {
                contractorSelect = value;
                OnPropertyChanged(nameof(ContractorSelect));
            }
        }
        public IEnumerable<DIC_Contractor> Contractors
        {
            get { return contractors; }
            set
            {
                contractors = value;
                OnPropertyChanged(nameof(Contractors));
            }
        }
        public decimal Percent
        {
            get { return percent; }
            set
            {
                percent = value;
                OnPropertyChanged(nameof(Percent));
            }
        }
        public decimal Adjust
        {
            get { return adjust; }
            set
            {
                adjust = value;
                OnPropertyChanged(nameof(Adjust));
            }
        }
        public string TaxSelect
        {
            get { return taxSelect; }
            set
            {
                taxSelect = value;
                OnPropertyChanged(nameof(TaxSelect));
            }
        }
        public List<string> Tax
        {
            get { return tax; }
            set
            {
                tax = value;
                OnPropertyChanged(nameof(Tax));
            }
        }
        public decimal AdjustContractor
        {
            get { return adjustContractor; }
            set
            {
                adjustContractor = value;
                OnPropertyChanged(nameof(AdjustContractor));
            }
        }

        private Command _insLabour;
        private Command _insContractor;

        public Command InsLabour => _insLabour ?? (_insLabour = new Command(obj=> 
        {
            if (LabourSelect != null)
            {
                LabourSelect.Contractor = ContractorSelect.Name;
                LabourSelect.Percent = Percent;
                LabourSelect.Adjust = Adjust;
                LabourSelect.Payout = decimal.Round(LabourSelect.Price * (Percent / 100m), 2);
                LabourSelect.Profit = decimal.Round((LabourSelect.Price - LabourSelect.Payout) + Adjust, 2);

                db.Entry(LabourSelect).State = EntityState.Modified;
                db.SaveChanges();

                Labours = null;
                Labours = db.Labours.Local.ToBindingList().Where(l => l.LabourProfitId == LabourProfit.Id);

                LoadLabourContractor();
            }
        }));
        public Command InsContractor => _insContractor ?? (_insContractor = new Command(obj=> 
        {
            if (LabourContractorSelect != null)
            {
                LabourContractorSelect.TAX = TaxSelect;
                LabourContractorSelect.Adjust = AdjustContractor;
                LabourContractorSelect.Total = decimal.Round(LabourContractorSelect.Payout + LabourContractorSelect.Adjust, 2);
                LabourContractorSelect.GST = (LabourContractorSelect.TAX == "Yes") ? (decimal.Round(LabourContractorSelect.Total * 0.05m, 2)) : (0m);
                LabourContractorSelect.TotalContractor = decimal.Round(LabourContractorSelect.Total + LabourContractorSelect.GST, 2);

                db.Entry(LabourContractorSelect).State = EntityState.Modified;
                db.SaveChanges();
                LabourContractors = null;
                LabourContractors = db.LabourContractors.Local.ToBindingList().Where(c => c.LabourProfitId == LabourProfit.Id).OrderBy(c=>c.Contractor);

                TaxSelect = null;
                AdjustContractor = 0m;

            }
        }));

        public LabourProfitViewModel(ref BuilderContext context, LabourProfit select)
        {
            db = context;
            db.Labours.Load();
            db.LabourContractors.Load();
            db.DIC_Contractors.Load();
            LabourProfit = select;
            Labours = db.Labours.Local.ToBindingList().Where(l=>l.LabourProfitId == LabourProfit.Id);
            LabourContractors = db.LabourContractors.Local.ToBindingList().Where(c=>c.LabourProfitId == LabourProfit.Id).OrderBy(c=>c.Contractor);
            Contractors = db.DIC_Contractors.Local.ToBindingList();
            TaxLoad();
        }
        private void TaxLoad()
        {
            Tax = new List<string>();
            Tax.Add("Yes");
            Tax.Add("No");
        }
        private void LoadLabourContractor()
        {
            var contract = Labours.Select(i => i.Contractor)?.Distinct()?.OrderBy(x => 1);
            db.LabourContractors.RemoveRange(LabourContractors);
            db.SaveChanges();
            if (contract != null)
            {                
                foreach (var item in contract)
                {
                    if (item != null)
                    {
                        decimal temp = decimal.Round(Labours.Where(i => i.Contractor == item).Select(i => i.Payout).Sum(), 2);
                        decimal total = decimal.Round(temp + AdjustContractor, 2);
                        decimal gst = (TaxSelect == "Yes") ? (decimal.Round((temp + AdjustContractor) * 0.05m, 2)) : (0m);

                        var test = LabourContractors.FirstOrDefault(c => c.Contractor == item);
                        if (test == null)
                        {
                            LabourContractor _Contractor = new LabourContractor()
                            {
                                Contractor = item,
                                Payout = temp,
                                Adjust = 0m,
                                TAX = TaxSelect,
                                Total = total,
                                GST = 0m,
                                TotalContractor = total,
                                LabourProfitId = LabourProfit.Id
                            };
                            db.LabourContractors.Add(_Contractor);
                            db.SaveChanges();
                        }
                        else
                        {
                            test.Payout = temp;
                            test.Total = decimal.Round(test.Payout + test.Adjust, 2);
                            test.GST = (test.TAX == "Yes") ? (decimal.Round(test.Total * 0.05m, 2)) : (0m);
                            test.TotalContractor = decimal.Round(test.Total + test.GST, 2);

                            db.Entry(test).State = EntityState.Modified;
                            db.SaveChanges();
                        }

                        LabourContractors = null;
                        LabourContractors = db.LabourContractors.Local.ToBindingList().Where(c => c.LabourProfitId == LabourProfit.Id).OrderBy(c => c.Contractor);
                    }
                }
            }

        }
    }
}
