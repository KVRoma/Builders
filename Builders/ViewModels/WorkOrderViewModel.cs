using Builders.Commands;
using Builders.Enums;
using Builders.Models;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.ViewModels
{
    public class WorkOrderViewModel : ViewModel
    {
        private BuilderContext db;
        public string WindowName { get; } = "Work Order";
        private bool enableWork;
        private bool enableButtonCreat;
        
        //****************************************************************** Models privat
        private WorkOrder orderSelect;
        private IEnumerable<WorkOrder> orders;
        private WorkOrder_Work workSelect;
        private IEnumerable<WorkOrder_Work> works;
        private IList<WorkOrder_Installation> installationSelect;    // TEST
        private IEnumerable<WorkOrder_Installation> installations;
        private WorkOrder_Contractor contractorSelect;
        private IEnumerable<WorkOrder_Contractor> contractors;
        //******************************************************************** ComboBox privat
        private Quotation quotationSelect;
        private IEnumerable<Quotation> quotations;
        private string trimSelect;
        private List<string> trim;
        private string baseboardSelect;
        private List<string> baseboard;
        private string colourSelect;
        private List<string> colour;
        private string replacingSelect;
        private List<string> replacing;
        private string taxSelect;
        private List<string> tax;
        private DIC_Area areaSelect;
        private IEnumerable<DIC_Area> areas;
        private DIC_Room roomSelect;
        private IEnumerable<DIC_Room> rooms;
        private DIC_ExistingFloor floorSelect;
        private IEnumerable<DIC_ExistingFloor> floors;
        private string furnitureSelect;
        private List<string> furnitures;
        private DIC_Contractor installerSelect;
        private IEnumerable<DIC_Contractor> installers;
        private DIC_Contractor contractorWorkSelect;
        private IEnumerable<DIC_Contractor> contractorWorks;
        //******************************************************************** TextBox privat
        private string parking;
        private DateTime serviceDate;
        private DateTime completionDate;
        private string lf;
        private int? pieces;
        private string newFloor;
        private string misc;
        private decimal percent;
        private decimal ajust;
        private decimal rate;

        public bool EnableWork
        {
            get { return enableWork; }
            set
            {
                enableWork = value;
                OnPropertyChanged(nameof(EnableWork));
            }
        }
        public bool EnableButtonCreat
        {
            get { return enableButtonCreat; }
            set
            {
                enableButtonCreat = value;
                OnPropertyChanged(nameof(EnableButtonCreat));
            }
        }
        //********************************************************************* Models public
        public WorkOrder OrderSelect
        {
            get { return orderSelect; }
            set
            {
                orderSelect = value;
                OnPropertyChanged(nameof(OrderSelect));
            }
        }
        public IEnumerable<WorkOrder> Orders
        {
            get { return orders; }
            set
            {
                orders = value;
                OnPropertyChanged(nameof(Orders));
            }
        }
        public WorkOrder_Work WorkSelect
        {
            get { return workSelect; }
            set
            {
                workSelect = value;
                OnPropertyChanged(nameof(WorkSelect));
                LoadComboBoxWorkSelect(WorkSelect);
            }
        }
        public IEnumerable<WorkOrder_Work> Works
        {
            get { return works; }
            set
            {
                works = value;
                OnPropertyChanged(nameof(Works));
            }
        }
        public IList<WorkOrder_Installation> InstallationSelect           
        {
            get { return installationSelect; }
            set
            {
                installationSelect = value;
                OnPropertyChanged(nameof(InstallationSelect));                              
            }
        }
        public IEnumerable<WorkOrder_Installation> Installations
        {
            get { return installations; }
            set
            {
                installations = value;
                OnPropertyChanged(nameof(Installations));
            }
        }
        public WorkOrder_Contractor ContractorSelect
        {
            get { return contractorSelect; }
            set
            {
                contractorSelect = value;
                OnPropertyChanged(nameof(ContractorSelect));
                if (ContractorSelect != null)
                {
                    TaxSelect = ContractorSelect.TAX;
                }
            }
        }
        public IEnumerable<WorkOrder_Contractor> Contractors
        {
            get { return contractors; }
            set
            {
                contractors = value;
                OnPropertyChanged(nameof(Contractors));
            }
        }
        //********************************************************************* ComboBox public
        public Quotation QuotationSelect
        {
            get { return quotationSelect; }
            set
            {
                quotationSelect = value;
                OnPropertyChanged(nameof(QuotationSelect));
                LoadInstallation(QuotationSelect);
            }
        }
        public IEnumerable<Quotation> Quotations
        {
            get { return quotations; }
            set
            {
                quotations = value;
                OnPropertyChanged(nameof(Quotations));
            }
        }
        public string TrimSelect
        {
            get { return trimSelect; }
            set
            {
                trimSelect = value;
                OnPropertyChanged(nameof(TrimSelect));
            }
        }
        public List<string> Trim
        {
            get { return trim; }
            set
            {
                trim = value;
                OnPropertyChanged(nameof(Trim));
            }
        }
        public string BaseboardSelect
        {
            get { return baseboardSelect; }
            set
            {
                baseboardSelect = value;
                OnPropertyChanged(nameof(BaseboardSelect));
            }
        }
        public List<string> Baseboard
        {
            get { return baseboard; }
            set
            {
                baseboard = value;
                OnPropertyChanged(nameof(Baseboard));
            }
        }
        public string ColourSelect
        {
            get { return colourSelect; }
            set
            {
                colourSelect = value;
                OnPropertyChanged(nameof(ColourSelect));
            }
        }
        public List<string> Colour
        {
            get { return colour; }
            set
            {
                colour = value;
                OnPropertyChanged(nameof(Colour));
            }
        }
        public string ReplacingSelect
        {
            get { return replacingSelect; }
            set
            {
                replacingSelect = value;
                OnPropertyChanged(nameof(replacingSelect));
            }
        }
        public List<string> Replacing
        {
            get { return replacing; }
            set
            {
                replacing = value;
                OnPropertyChanged(nameof(Replacing));
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
        public DIC_Area AreaSelect
        {
            get { return areaSelect; }
            set
            {
                areaSelect = value;
                OnPropertyChanged(nameof(AreaSelect));
            }
        }
        public IEnumerable<DIC_Area> Areas
        {
            get { return areas; }
            set
            {
                areas = value;
                OnPropertyChanged(nameof(Areas));
            }
        }
        public DIC_Room RoomSelect
        {
            get { return roomSelect; }
            set
            {
                roomSelect = value;
                OnPropertyChanged(nameof(RoomSelect));
            }
        }
        public IEnumerable<DIC_Room> Rooms
        {
            get { return rooms; }
            set
            {
                rooms = value;
                OnPropertyChanged(nameof(Rooms));
            }
        }
        public DIC_ExistingFloor FloorSelect
        {
            get { return floorSelect; }
            set
            {
                floorSelect = value;
                OnPropertyChanged(nameof(FloorSelect));
            }
        }
        public IEnumerable<DIC_ExistingFloor> Floors
        {
            get { return floors; }
            set
            {
                floors = value;
                OnPropertyChanged(nameof(Floors));
            }
        }
        public string FurnitureSelect
        {
            get { return furnitureSelect; }
            set
            {
                furnitureSelect = value;
                OnPropertyChanged(nameof(FurnitureSelect));
            }
        }
        public List<string> Furnitures
        {
            get { return furnitures; }
            set
            {
                furnitures = value;
                OnPropertyChanged(nameof(Furnitures));
            }
        }
        public DIC_Contractor InstallerSelect
        {
            get { return installerSelect; }
            set
            {
                installerSelect = value;
                OnPropertyChanged(nameof(InstallerSelect));
                if (InstallerSelect?.NameAndPhone == "")
                {
                    Percent = 0;
                    OnPropertyChanged(nameof(Percent));
                }
            }
        }
        public IEnumerable<DIC_Contractor> Installers
        {
            get { return installers; }
            set 
            {
                installers = value;
                OnPropertyChanged(nameof(Installers));
            }
        }
        public DIC_Contractor ContractorWorkSelect
        {
            get { return contractorWorkSelect; }
            set
            {
                contractorWorkSelect = value;
                OnPropertyChanged(nameof(ContractorWorkSelect));
            }
        }
        public IEnumerable<DIC_Contractor> ContractorWorks
        {
            get { return contractorWorks; }
            set
            {
                contractorWorks = value;
                OnPropertyChanged(nameof(ContractorWorks));
            }
        }

        //******************************************************************** TextBox public
        public string Parking
        {
            get { return parking; }
            set
            {
                parking = value;
                OnPropertyChanged(nameof(Parking));
            }
        }
        public DateTime ServiceDate
        {
            get { return serviceDate; }
            set
            {
                serviceDate = value;
                OnPropertyChanged(nameof(ServiceDate));
            }
        }
        public DateTime CompletionDate
        {
            get { return completionDate; }
            set
            {
                completionDate = value;
                OnPropertyChanged(nameof(CompletionDate));
            }
        }
        public string LF
        {
            get { return lf; }
            set
            {
                lf = value;
                OnPropertyChanged(nameof(LF));
            }
        }
        public int? Pieces
        {
            get { return pieces; }
            set
            {
                pieces = value;
                OnPropertyChanged(nameof(Pieces));
            }
        }
        public string NewFloor
        {
            get { return newFloor; }
            set
            {
                newFloor = value;
                OnPropertyChanged(nameof(NewFloor));
            }
        }
        public string Misc
        {
            get { return misc; }
            set
            {
                misc = value;
                OnPropertyChanged(nameof(Misc));
            }
        }
        public decimal Percent
        {
            get { return percent; }
            set 
            {
                percent = value;
                OnPropertyChanged(nameof(Percent));
                if (Percent != 0)
                {
                    Rate = 0m;
                }
            }
                
        }
        public decimal Ajust
        {
            get { return ajust; }
            set
            {
                ajust = value;
                OnPropertyChanged(nameof(Ajust));
            }
        }
        public decimal Rate
        {
            get { return rate; }
            set
            {
                rate = value;
                OnPropertyChanged(nameof(Rate));
                if (Rate != 0m)
                {
                    Percent = 0;
                }
            }
        }
        //********************************************************************* Command privat
        private Command _creatOrder;
        private Command _addWork;
        private Command _insWork;
        private Command _delWork;
        private Command _addInstallation;
        private Command _addContractor;
        //********************************************************************* Command public
        public Command CreatOrder => _creatOrder ?? (_creatOrder = new Command(obj=> 
        {
            if (QuotationSelect != null)
            {
                WorkOrder workorder = new WorkOrder() 
                {
                    QuotaId = QuotationSelect.Id,
                    NumberQuota = QuotationSelect.NumberQuota,
                    FirstName = QuotationSelect.FirstName,
                    LastName = QuotationSelect.LastName,
                    PhoneNumber = QuotationSelect.PhoneNumber,
                    Email = QuotationSelect.Email,
                    DateWork = DateTime.Today,
                    Parking = Parking,
                    DateServices = ServiceDate,
                    DateCompletion = CompletionDate,
                    Trim = TrimSelect,
                    Colour = ColourSelect,
                    LF = LF,
                    Baseboard = BaseboardSelect,
                    ReplacingYesNo = ReplacingSelect,
                    ReplacingQuantity = Pieces,                    
                };
                db.WorkOrders.Add(workorder);
                db.SaveChanges();

                db.WorkOrders.Load();
                OrderSelect = db.WorkOrders.Local.ToBindingList().OrderByDescending(q => q.Id).FirstOrDefault();

                foreach (var item in Installations)
                {
                    item.WorkOrderId = OrderSelect.Id;
                    db.WorkOrder_Installations.Add(item);
                    db.SaveChanges();
                }

                Installations = db.WorkOrder_Installations.Local.ToBindingList().Where(i => i.WorkOrderId == OrderSelect.Id).OrderBy(i => i.Groupe).ThenBy(c => c.Contractor);
                EnableWork = true;
                EnableButtonCreat = false;
            }
        }));
        public Command AddWork => _addWork ?? (_addWork = new Command(obj=> 
        {
            WorkOrder_Work work = new WorkOrder_Work() 
            {
                Area = AreaSelect?.Name,
                Room = RoomSelect?.Name,
                Existing = FloorSelect?.Name,
                NewFloor = NewFloor,
                Furniture = FurnitureSelect,
                Misc = Misc,
                Contractor = ContractorWorkSelect?.Name,
                Color = ContractorWorkSelect?.Color,
                WorkOrderId = OrderSelect.Id
            };
            db.WorkOrder_Works.Add(work);
            db.SaveChanges();
            Works = db.WorkOrder_Works.Local.ToBindingList().Where(w => w.WorkOrderId == OrderSelect.Id);            
        }));
        public Command InsWork => _insWork ?? (_insWork = new Command(obj=> 
        {
            if (WorkSelect != null)
            {
                WorkSelect.Area = AreaSelect?.Name;
                WorkSelect.Room = RoomSelect?.Name;
                WorkSelect.Existing = FloorSelect?.Name;
                WorkSelect.NewFloor = NewFloor;
                WorkSelect.Furniture = FurnitureSelect;
                WorkSelect.Misc = Misc;
                WorkSelect.Contractor = ContractorWorkSelect?.Name;
                WorkSelect.Color = ContractorWorkSelect?.Color;
                db.Entry(WorkSelect).State = EntityState.Modified;
                db.SaveChanges();
                Works = null;
                Works = db.WorkOrder_Works.Local.ToBindingList().Where(w=>w.WorkOrderId == OrderSelect.Id);
            }
        }));
        public Command DelWork => _delWork ?? (_delWork = new Command(obj=> 
        {
            if (WorkSelect != null)
            {
                db.WorkOrder_Works.Remove(WorkSelect);
                db.SaveChanges();
                Works = null;
                Works = db.WorkOrder_Works.Local.ToBindingList().Where(w => w.WorkOrderId == OrderSelect.Id);
            }
        }));
        public Command AddInstallation => _addInstallation ?? (_addInstallation = new Command(obj=> 
        {
           
            if (InstallationSelect != null)                         
            {
                foreach (var item in InstallationSelect)              
                {
                    item.Contractor = InstallerSelect?.Name;
                    item.Color = InstallerSelect?.Color;
                    item.Procent = Percent;
                    if (Rate != 0m)
                    {
                        //item.Rate = Rate;
                        //item.Price = decimal.Round(item.Quantity * item.Rate, 2);
                        //item.Payout = item.Price;
                        item.Procent = Rate;
                        item.Payout = decimal.Round(item.Quantity * item.Procent, 2);
                    }
                    else
                    {
                        item.Payout = decimal.Round(item.Price * (item.Procent / 100m), 2);
                    }
                    db.Entry(item).State = EntityState.Modified;
                    db.SaveChanges();
                }
                Installations = null;
                Installations = db.WorkOrder_Installations.Local.ToBindingList().Where(i=>i.WorkOrderId == OrderSelect.Id).OrderBy(i=>i.Groupe);   //.ThenBy(c => c.Contractor);

                LoadContractor();
            }
        }));
        public Command AddContractor => _addContractor ?? (_addContractor = new Command(obj=> 
        {
            if (ContractorSelect != null)
            {
                ContractorSelect.TAX = TaxSelect;
                ContractorSelect.Adjust = Ajust;
                ContractorSelect.Total = decimal.Round(ContractorSelect.Payout + ContractorSelect.Adjust, 2);
                ContractorSelect.GST = (ContractorSelect.TAX == "Yes") ? (decimal.Round(ContractorSelect.Total * 0.05m, 2)) : (0m);
                ContractorSelect.TotalContractor = decimal.Round(ContractorSelect.Total + ContractorSelect.GST, 2);

                db.Entry(ContractorSelect).State = EntityState.Modified;
                db.SaveChanges();
                Contractors = null;
                Contractors = db.WorkOrder_Contractors.Local.ToBindingList().Where(c=>c.WorkOrderId == OrderSelect.Id).OrderBy(c=>c.Contractor);
                                
                Ajust = 0m;
            }
        }));

       
        //*********************************************************************
        public WorkOrderViewModel(ref BuilderContext context, EnumClient client, int? IdWorkOrder)
        {
            db = context;
            
            db.WorkOrder_Works.Load();
            db.WorkOrder_Installations.Load();
            db.WorkOrder_Contractors.Load();
            db.DIC_Areas.Load();
            db.DIC_Rooms.Load();
            db.DIC_ExistingFloors.Load();
            db.DIC_Contractors.Load();

            Works = db.WorkOrder_Works.Local.ToBindingList();
            Installations = db.WorkOrder_Installations.Local.ToBindingList().OrderBy(i => i.Groupe);                            //.ThenBy(c => c.Contractor);
            InstallationSelect = new List<WorkOrder_Installation>();
            Contractors = db.WorkOrder_Contractors.Local.ToBindingList();
            Areas = db.DIC_Areas.Local.ToBindingList().OrderBy(a=>a.Name);
            Rooms = db.DIC_Rooms.Local.ToBindingList().OrderBy(a => a.Name);
            Floors = db.DIC_ExistingFloors.Local.ToBindingList().OrderBy(a => a.Name);
            Installers = db.DIC_Contractors.Local.ToBindingList().OrderBy(a => a.Name);
            ContractorWorks = db.DIC_Contractors.Local.ToBindingList().OrderBy(a => a.Name);

            TrimLoad();
            BaseboardLoad();
            //ColourLoad();
            ReplacingLoad();
            TaxLoad();
            FurnitureLoad();

            switch (client)
            {
                case EnumClient.Add:
                    {
                        ServiceDate = DateTime.Today;
                        CompletionDate = DateTime.Today;
                        EnableWork = false;
                        EnableButtonCreat = true;
                        Works = db.WorkOrder_Works.Local.ToBindingList().Where(w => w.WorkOrderId == OrderSelect?.Id);
                        Contractors = db.WorkOrder_Contractors.Local.ToBindingList().Where(c => c.WorkOrderId == OrderSelect?.Id); 
                    }
                    break;
                case EnumClient.Ins:
                    {
                        OrderSelect = db.WorkOrders.FirstOrDefault(o=>o.Id == IdWorkOrder);
                        Parking = OrderSelect?.Parking;
                        ServiceDate = OrderSelect.DateServices;
                        CompletionDate = OrderSelect.DateCompletion;
                        TrimSelect = OrderSelect.Trim;
                        ColourSelect = OrderSelect.Colour;
                        LF = OrderSelect.LF;
                        BaseboardSelect = OrderSelect.Baseboard;
                        ReplacingSelect = OrderSelect.ReplacingYesNo;
                        Pieces = OrderSelect.ReplacingQuantity;
                        QuotationSelect = db.Quotations.FirstOrDefault(q => q.Id == OrderSelect.QuotaId);
                        Works = db.WorkOrder_Works.Local.ToBindingList().Where(w => w.WorkOrderId == OrderSelect.Id);
                        Installations = db.WorkOrder_Installations.Local.ToBindingList().Where(i => i.WorkOrderId == OrderSelect.Id).OrderBy(i => i.Groupe).ThenBy(c => c.Contractor);
                        Contractors = db.WorkOrder_Contractors.Local.ToBindingList().Where(c => c.WorkOrderId == OrderSelect.Id).OrderBy(c=>c.Contractor);
                        EnableWork = true;
                        EnableButtonCreat = false;
                    }
                    break;
            }
            
        }
        
        //********************************************************************

        private void TrimLoad()
        {
            Trim = new List<string>();
            Trim.Add("None");
            Trim.Add("Window");
            Trim.Add("Kitchen Cabinetry");
            Trim.Add("Window & Kitchen Cabinetry");
        }
        private void BaseboardLoad()
        {
            Baseboard = new List<string>();
            Baseboard.Add("Remove & Reuse");
            Baseboard.Add("Remove & Dispose");
        }
        //private void ColourLoad()
        //{
        //    Colour = new List<string>();
        //    Colour.Add(nameof(Color.Blue));
        //    Colour.Add(nameof(Color.Red));
        //    Colour.Add(nameof(Color.Green));
        //}
        private void ReplacingLoad()
        {
            Replacing = new List<string>();
            Replacing.Add("Yes");
            Replacing.Add("No");
        }
        private void TaxLoad()
        {
            Tax = new List<string>();
            Tax.Add("Yes");
            Tax.Add("No");
        }
        private void FurnitureLoad()
        {
            Furnitures = new List<string>();
            Furnitures.Add("Yes");
            Furnitures.Add("No");
        }

        private void LoadInstallation(Quotation quota)
        {
            if (quota != null)
            {
                List<WorkOrder_Installation> ins = new List<WorkOrder_Installation>();
                var material = db.MaterialQuotations.Where(m => m.QuotationId == quota.Id)?.OrderBy(m => m.Groupe);
                if (material != null)
                {
                    foreach (MaterialQuotation item in material)
                    {
                        if ((item.Groupe == "INSTALLATION") || (item.Groupe == "DEMOLITION") || (item.Groupe == "OPTIONAL SERVICES") || (item.Groupe == "FLOORING DELIVERY"))
                        {
                            ins.Add(new WorkOrder_Installation()
                            {
                                Groupe = item.Groupe,
                                Item = item.Item,
                                Description = item.Description,
                                Quantity = item.Quantity,
                                Rate = item.Rate,
                                Price = item.Price,
                                Procent = 0,
                                Payout = 0,
                                Color = "Black"
                            });
                        }
                    }
                    Installations = null;
                    Installations = ins.OrderBy(i => i.Groupe); //.ThenBy(c => c.Contractor);                    
                }
            }
        }
        private void LoadComboBoxWorkSelect(WorkOrder_Work select)
        {
            if (select != null)
            {
                AreaSelect = db.DIC_Areas.FirstOrDefault(a => a.Name == select.Area);
                RoomSelect = db.DIC_Rooms.FirstOrDefault(r => r.Name == select.Room);
                FloorSelect = db.DIC_ExistingFloors.FirstOrDefault(f => f.Name == select.Existing);
                NewFloor = select.NewFloor;
                FurnitureSelect = select.Furniture;
                Misc = select.Misc;
            }
        }
        private void LoadContractor()
        {
            var contract = Installations.Select(i=>i.Contractor)?.Distinct()?.OrderBy(x=>1);
            db.WorkOrder_Contractors.RemoveRange(Contractors);
            db.SaveChanges();
            Contractors = null;
            Contractors = db.WorkOrder_Contractors.Local.ToBindingList().Where(c => c.WorkOrderId == OrderSelect.Id).OrderBy(c => c.Contractor);
            if (contract != null)
            {
                foreach (var item in contract)
                {
                    if (item != null)
                    {
                        decimal temp = decimal.Round(Installations.Where(i => i.Contractor == item).Select(i => i.Payout).Sum(), 2);
                        decimal total = decimal.Round(temp + Ajust, 2);
                        decimal gst = (TaxSelect == "Yes") ? (decimal.Round((temp + Ajust) * 0.05m, 2)) : (0m);

                        var test = Contractors.FirstOrDefault(c=>c.Contractor == item);
                        if (test == null)
                        {                            
                            WorkOrder_Contractor _Contractor = new WorkOrder_Contractor()
                            {
                                Contractor = item,
                                Color = db.DIC_Contractors.FirstOrDefault(c=>c.Name == item)?.Color,
                                Payout = temp,
                                Adjust = 0m,
                                TAX = Tax[0],  // Default "Yes"
                                Total = total,
                                GST = decimal.Round(total * 0.05m, 2), // Default "5%",
                                TotalContractor = decimal.Round(total + decimal.Round(total * 0.05m, 2), 2),
                                WorkOrderId = OrderSelect.Id
                            };
                            db.WorkOrder_Contractors.Add(_Contractor);
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

                        Contractors = null;
                        Contractors = db.WorkOrder_Contractors.Local.ToBindingList().Where(c=>c.WorkOrderId == OrderSelect.Id).OrderBy(c=>c.Contractor);
                    }
                }
            }

        }
        private void SaveEditWorkOrder(WorkOrder work)
        {
            work.Parking = Parking;
            work.DateServices = ServiceDate;
            work.DateCompletion = CompletionDate;
            work.Trim = TrimSelect;
            work.Colour = ColourSelect;
            work.LF = LF;
            work.Baseboard = BaseboardSelect;
            work.ReplacingYesNo = ReplacingSelect;
            work.ReplacingQuantity = Pieces;
            db.Entry(work).State = EntityState.Modified;
            db.SaveChanges();
        }
    }
}
