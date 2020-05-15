using Builders.Commands;
using Builders.Enums;
using Builders.Models;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Builders.ViewModels
{
    public class QuotationViewModel : ViewModel
    {
        public string WindowName { get; } = "Builder - Quotation";
        public BuilderContext db;
        public int? QuotaId;
        private EnumClient result;
        private bool flagCreatQuota;
       

        private Quotation quotaSelect;
        private IEnumerable<Quotation> quotas;

        private MaterialQuotation materialQuotationSelect;
        private IEnumerable<MaterialQuotation> materialQuotations;

        private IEnumerable<MaterialQuotation> quotationsLabour;
        private IEnumerable<MaterialQuotation> quotationsMaterial;

        private DIC_GroupeQuotation dic_GroupeSelect;
        private IEnumerable<DIC_GroupeQuotation> dic_Groupes;

        private DIC_ItemQuotation dic_ItemSelect;
        private IEnumerable<DIC_ItemQuotation> dic_Items;

        private DIC_DescriptionQuotation dic_DescriptionSelect;
        private IEnumerable<DIC_DescriptionQuotation> dic_Descriptions;

        private string dic_SectionSelect;
        private List<string> dic_Section;

        private Client clientSelect;
        private IEnumerable<Client> clients;
        private bool copyClient; // для сопіювання квоти на другого клієнта
        private int? lastIdClient;               // Для збереження попереднього клієнта  

        private decimal quantity;
        private decimal rate;
        //************************************************************************************************
        
        public Quotation QuotaSelect
        {
            get { return quotaSelect; }
            set
            {
                quotaSelect = value;
                OnPropertyChanged(nameof(QuotaSelect));
                MaterialQuotations = (QuotaSelect != null) ? (db.MaterialQuotations.Local.ToBindingList().Where(m => m.QuotationId == QuotaSelect.Id)) : null;                
            }
        }
        public IEnumerable<Quotation> Quotas
        {
            get { return quotas; }
            set
            {
                quotas = value;
                OnPropertyChanged(nameof(Quotas));
            }
        }

        public MaterialQuotation MaterialQuotationSelect
        {
            get {return materialQuotationSelect; }
            set 
            { 
                materialQuotationSelect = value; 
                OnPropertyChanged(nameof(MaterialQuotationSelect));
                
                if (MaterialQuotationSelect != null)
                {

                    if (MaterialQuotationSelect.Groupe == "FLOORING" || MaterialQuotationSelect.Groupe == "ACCESSORIES")
                    {
                        DIC_SectionSelect = "Material";
                    }
                    else
                    {
                        DIC_SectionSelect = "Labour";
                    }

                    LoadGroup();

                    DIC_GroupeSelect = (MaterialQuotationSelect.Groupe != null) ? DIC_Groupes.FirstOrDefault(g => g.NameGroupe == MaterialQuotationSelect.Groupe) : null;
                    DIC_ItemSelect = (MaterialQuotationSelect.Item != null) ? DIC_Items.FirstOrDefault(i => i.Name == MaterialQuotationSelect.Item) : null;
                    DIC_DescriptionSelect = (MaterialQuotationSelect.Description != null) ? DIC_Descriptions.FirstOrDefault(d => d.Name == MaterialQuotationSelect.Description) : null;
                    Quantity = MaterialQuotationSelect.Quantity;
                    Rate = MaterialQuotationSelect.Rate;
                }
            }
        }
        public IEnumerable<MaterialQuotation> MaterialQuotations 
        {
            get {return materialQuotations; } 
            set 
            { 
                materialQuotations = value; 
                OnPropertyChanged(nameof(MaterialQuotations));
                MaterialView();
            }
        }
        
        public IEnumerable<MaterialQuotation> QuotationsLabour
        {
            get { return quotationsLabour; }
            set
            {
                quotationsLabour = value;
                OnPropertyChanged(nameof(QuotationsLabour));
            }
        }
        public IEnumerable<MaterialQuotation> QuotationsMaterial
        {
            get { return quotationsMaterial; }
            set
            {
                quotationsMaterial = value;
                OnPropertyChanged(nameof(QuotationsMaterial));
            }
        }

        public DIC_GroupeQuotation DIC_GroupeSelect 
        { 
            get => dic_GroupeSelect; 
            set 
            { 
                dic_GroupeSelect = value; 
                OnPropertyChanged(nameof(DIC_GroupeSelect));
                DIC_Items = (DIC_GroupeSelect != null) ? db.DIC_ItemQuotations.Local.ToBindingList().Where(i => i.GroupeId == DIC_GroupeSelect.Id).OrderBy(i=>i.Name) : null;
            }
        }
        public IEnumerable<DIC_GroupeQuotation> DIC_Groupes 
        { 
            get => dic_Groupes; 
            set 
            { 
                dic_Groupes = value; 
                OnPropertyChanged(nameof(DIC_Groupes)); 
            }
        }
        
        public DIC_ItemQuotation DIC_ItemSelect 
        { 
            get => dic_ItemSelect; 
            set 
            {
                dic_ItemSelect = value; 
                OnPropertyChanged(nameof(DIC_ItemSelect));
                DIC_Descriptions = (DIC_ItemSelect != null) ? (db.DIC_DescriptionQuotations.Local.ToBindingList().Where(d => d.ItemId == DIC_ItemSelect.Id).OrderBy(i => i.Name)) : null;
            }
        }
        public IEnumerable<DIC_ItemQuotation> DIC_Items 
        { 
            get => dic_Items; 
            set 
            { 
                dic_Items = value; 
                OnPropertyChanged(nameof(DIC_Items)); 
            }
        }
        
        public DIC_DescriptionQuotation DIC_DescriptionSelect 
        { 
            get => dic_DescriptionSelect; 
            set 
            {
                dic_DescriptionSelect = value; 
                OnPropertyChanged(nameof(DIC_DescriptionSelect));
                Rate = (DIC_DescriptionSelect != null) ? (DIC_DescriptionSelect.Price) : 0;
            }
        }
        public IEnumerable<DIC_DescriptionQuotation> DIC_Descriptions 
        { 
            get => dic_Descriptions; 
            set 
            {
                dic_Descriptions = value; 
                OnPropertyChanged(nameof(DIC_Descriptions)); 
            }
        }

        public string DIC_SectionSelect
        {
            get { return dic_SectionSelect; }
            set
            {
                dic_SectionSelect = value;
                OnPropertyChanged(nameof(DIC_SectionSelect));

                LoadGroup();
            }
        }
        public List<string> DIC_Section
        {
            get { return dic_Section; }
            set
            {
                dic_Section = value;
                OnPropertyChanged(nameof(DIC_Section));
            }
        }

        public Client ClientSelect 
        { 
            get => clientSelect;
            set 
            { 
                clientSelect = value; 
                OnPropertyChanged(nameof(ClientSelect));
                if (ClientSelect != null)
                {
                    if (ClientSelect.Id != lastIdClient)
                    {
                        CopyClient = true;
                    }
                    else if (ClientSelect.Id == lastIdClient)
                    {
                        CopyClient = false;
                    }
                }
            } 
        }
        public IEnumerable<Client> Clients
        {
            get => clients;
            set 
            { 
                clients = value; 
                OnPropertyChanged(nameof(Clients)); 
            }
        }
        public bool CopyClient
        {
            get { return copyClient; }
            set
            {
                copyClient = value;
                OnPropertyChanged(nameof(copyClient));
            }
        }

        public decimal Quantity
        {
            get => quantity;
            set 
            { 
                quantity = value; 
                OnPropertyChanged(nameof(Quantity)); 
            }
        }
        public decimal Rate
        {
            get => rate;
            set 
            { 
                rate = value; 
                OnPropertyChanged(nameof(Rate)); 
            }
        }
        //**************************************************************************************************
        private Command _addItem;
        private Command _insItem;
        private Command _delItem;
        private Command _otherQuotation;        
        

        public Command AddItem => _addItem ?? (_addItem = new Command(obj => 
        {
            if (ClientSelect != null && flagCreatQuota == false)
            {
                Quotation quotation = new Quotation()
                {
                    
                    QuotaDate = DateTime.Today,
                    PrefixNumberQuota = "Q",
                    ClientId = ClientSelect.Id,
                    NumberClient = ClientSelect.NumberClient,
                    FirstName = ClientSelect.PrimaryFirstName,
                    LastName = ClientSelect.PrimaryLastName,
                    PhoneNumber = ClientSelect.PrimaryPhoneNumber,
                    Email = ClientSelect.PrimaryEmail,
                    JobDescription = "",
                    JobNote = "",

                    MaterialSubtotal = 0m,
                    //MaterialTax = 0m,
                    MaterialDiscountYN = 0m,
                    MaterialDiscountAmount = 0m,
                    //MaterialTotal = 0m,

                    LabourSubtotal = 0m,
                    //LabourTax = 0m,
                    LabourDiscountYN = 0m,
                    LabourDiscountAmount = 0m,
                    //LabourTotal = 0m,

                    //ProjectTotal = 0m,
                    FinancingYesNo = false,
                    //FinancingAmount = 0m,
                    //FinancingFee = 0m,
                    AmountPaidByCreditCard = 0m,
                    //ProcessingFee = 0m,
                    //InvoiceGrandTotal = 0m                    
                    
                };
                db.Quotations.Add(quotation);
                db.SaveChanges();
                flagCreatQuota = true;
                QuotaSelect = db.Quotations.Local.ToBindingList().OrderByDescending(q => q.Id).FirstOrDefault();
                QuotaId = QuotaSelect.Id;
            }
            if (flagCreatQuota && QuotaSelect != null)
            {
                
                var temp = db.MaterialQuotations.Where(f => f.QuotationId == QuotaSelect.Id);

                int flooring = (temp != null) ? (temp.Where(t=>t.Groupe == "FLOORING").Count()) : 0;
                int accessories = (temp != null) ? (temp.Where(t => t.Groupe == "ACCESSORIES").Count()) : 0;
                int installation = (temp != null) ? (temp.Where(t => t.Groupe == "INSTALLATION").Count()) : 0;
                int demolition = (temp != null) ? (temp.Where(t => t.Groupe == "DEMOLITION").Count()) : 0;
                int services = (temp != null) ? (temp.Where(t => t.Groupe == "OPTIONAL SERVICES").Count()) : 0;
                int delivery = (temp != null) ? (temp.Where(t => t.Groupe == "FLOORING DELIVERY").Count()) : 0;

                bool flagSaveItem = false;

                switch (DIC_GroupeSelect?.NameGroupe)
                {
                    case "FLOORING":
                        {
                            if (flooring > 7)
                            {
                                MessageBox.Show("The limit is reached...\n FLOORING (8 pcs.) !!!", "Warning !!!");
                            }
                            else
                            {
                                flagSaveItem = true;
                            }
                        }
                        break;
                    case "ACCESSORIES":
                        {
                            if (accessories > 15)
                            {
                                MessageBox.Show("The limit is reached...\n ACCESSORIES (16 pcs.) !!!", "Warning !!!");
                            }
                            else
                            {
                                flagSaveItem = true;
                            }
                        }
                        break;
                    case "INSTALLATION":
                        {
                            if (installation > 7)
                            {
                                MessageBox.Show("The limit is reached...\n INSTALLATION (8 pcs.) !!!", "Warning !!!");
                            }
                            else
                            {
                                flagSaveItem = true;
                            }
                        }
                        break;
                    case "DEMOLITION":
                        {
                            if (demolition > 4)
                            {
                                MessageBox.Show("The limit is reached...\n DEMOLITION (5 pcs.) !!!", "Warning !!!");
                            }
                            else
                            {
                                flagSaveItem = true;
                            }
                        }
                        break;
                    case "OPTIONAL SERVICES":
                        {
                            if (services > 6)
                            {
                                MessageBox.Show("The limit is reached...\n OPTIONAL SERVICES (7 pcs.) !!!", "Warning !!!");
                            }
                            else
                            {
                                flagSaveItem = true;
                            }
                        }
                        break;
                    case "FLOORING DELIVERY":
                        {
                            if (delivery > 3)
                            {
                                MessageBox.Show("The limit is reached...\n FLOORING DELIVERY (4 pcs.) !!!", "Warning !!!");
                            }
                            else
                            {
                                flagSaveItem = true;
                            }
                        }
                        break;
                    default:
                        { }
                        break;
                }
                
                if (flagSaveItem)
                {
                    MaterialQuotation material = new MaterialQuotation()
                    {
                        QuotationId = QuotaSelect.Id,
                        Groupe = DIC_GroupeSelect?.NameGroupe,
                        Item = DIC_ItemSelect?.Name,
                        Description = DIC_DescriptionSelect?.Name,
                        Rate = decimal.Round(Rate, 2),
                        Quantity = Quantity,
                        Price = decimal.Round(Rate * Quantity, 2),
                        SupplierId = DIC_ItemSelect?.SupplierId
                    };
                    db.MaterialQuotations.Add(material);
                    db.SaveChanges();

                    MaterialQuotations = null;
                    MaterialQuotations = db.MaterialQuotations.Local.ToBindingList().Where(m => m.QuotationId == QuotaSelect.Id);
                    MaterialQuotationSelect = material;
                }
            }
        }));
        public Command InsItem => _insItem ?? (_insItem = new Command(obj =>
        {
            if (MaterialQuotationSelect != null)
            {
                var temp = db.MaterialQuotations.Where(f => f.QuotationId == QuotaSelect.Id);
                var tempMaterialSelect = db.MaterialQuotations.FirstOrDefault(m => m.Id == MaterialQuotationSelect.Id);

                int flooring = (temp != null) ? (temp.Where(t => t.Groupe == "FLOORING").Count()) : 0;
                int accessories = (temp != null) ? (temp.Where(t => t.Groupe == "ACCESSORIES").Count()) : 0;
                int installation = (temp != null) ? (temp.Where(t => t.Groupe == "INSTALLATION").Count()) : 0;
                int demolition = (temp != null) ? (temp.Where(t => t.Groupe == "DEMOLITION").Count()) : 0;
                int services = (temp != null) ? (temp.Where(t => t.Groupe == "OPTIONAL SERVICES").Count()) : 0;
                int delivery = (temp != null) ? (temp.Where(t => t.Groupe == "FLOORING DELIVERY").Count()) : 0;

                bool flagSaveItem = false;

                switch (DIC_GroupeSelect?.NameGroupe)
                {
                    case "FLOORING":
                        {
                            if (flooring > 7)
                            {
                                MessageBox.Show("The limit is reached...\n FLOORING (8 pcs.) !!!", "Warning !!!");
                            }
                            else
                            {
                                flagSaveItem = true;
                            }
                        }
                        break;
                    case "ACCESSORIES":
                        {
                            if (accessories > 15)
                            {
                                MessageBox.Show("The limit is reached...\n ACCESSORIES (16 pcs.) !!!", "Warning !!!");
                            }
                            else
                            {
                                flagSaveItem = true;
                            }
                        }
                        break;
                    case "INSTALLATION":
                        {
                            if (installation > 7)
                            {
                                MessageBox.Show("The limit is reached...\n INSTALLATION (8 pcs.) !!!", "Warning !!!");
                            }
                            else
                            {
                                flagSaveItem = true;
                            }
                        }
                        break;
                    case "DEMOLITION":
                        {
                            if (demolition > 4)
                            {
                                MessageBox.Show("The limit is reached...\n DEMOLITION (5 pcs.) !!!", "Warning !!!");
                            }
                            else
                            {
                                flagSaveItem = true;
                            }
                        }
                        break;
                    case "OPTIONAL SERVICES":
                        {
                            if (services > 6)
                            {
                                MessageBox.Show("The limit is reached...\n OPTIONAL SERVICES (7 pcs.) !!!", "Warning !!!");
                            }
                            else
                            {
                                flagSaveItem = true;
                            }
                        }
                        break;
                    case "FLOORING DELIVERY":
                        {
                            if (delivery > 3)
                            {
                                MessageBox.Show("The limit is reached...\n FLOORING DELIVERY (4 pcs.) !!!", "Warning !!!");
                            }
                            else
                            {
                                flagSaveItem = true;
                            }
                        }
                        break;
                    default:
                        { }
                        break;
                }
                if (flagSaveItem)
                {
                    tempMaterialSelect.Groupe = DIC_GroupeSelect?.NameGroupe;
                    tempMaterialSelect.Item = DIC_ItemSelect?.Name;
                    tempMaterialSelect.SupplierId = DIC_ItemSelect?.SupplierId;
                    tempMaterialSelect.Description = DIC_DescriptionSelect?.Name;
                    tempMaterialSelect.Rate = decimal.Round(Rate, 2);
                    tempMaterialSelect.Quantity = Quantity;
                    tempMaterialSelect.Price = decimal.Round(Rate * Quantity, 2);

                    db.Entry(tempMaterialSelect).State = EntityState.Modified;
                    db.SaveChanges();
                    MaterialQuotationSelect = null;
                    MaterialQuotations = (QuotaSelect != null) ? (db.MaterialQuotations.Local.ToBindingList().Where(m => m.QuotationId == QuotaSelect.Id)) : null;
                }
            }
        }));
        public Command DelItem => _delItem ?? (_delItem = new Command(obj =>
        {
            if (MaterialQuotationSelect != null)
            {
                var temp = db.MaterialQuotations.FirstOrDefault(m => m.Id == MaterialQuotationSelect.Id);
                db.MaterialQuotations.Remove(temp);
                db.SaveChanges();
                MaterialQuotations = (QuotaSelect != null) ? (db.MaterialQuotations.Local.ToBindingList().Where(m => m.QuotationId == QuotaSelect.Id)) : null;
            }
        }));
        public Command OtherQuotation => _otherQuotation ?? (_otherQuotation = new Command(async obj =>
        {
            if (flagCreatQuota && QuotaSelect != null)
            {
                var displayRootRegistry = (Application.Current as App).displayRootRegistry;
                var other = new QuotaOtherViewModel()
                {
                    Description = QuotaSelect.JobDescription,
                    Notes = QuotaSelect.JobNote,
                    DiscountMaterial = QuotaSelect.MaterialDiscountAmount,
                    DiscountLabour = QuotaSelect.LabourDiscountAmount,
                    Financing = QuotaSelect.FinancingYesNo,
                    CreditCard = QuotaSelect.AmountPaidByCreditCard
                };
                await displayRootRegistry.ShowModalPresentation(other);
                if (other.PressOk)
                {
                    QuotaSelect.JobDescription = other.Description;
                    QuotaSelect.JobNote = other.Notes;
                    QuotaSelect.MaterialDiscountAmount = other.DiscountMaterial;
                    QuotaSelect.LabourDiscountAmount = other.DiscountLabour;
                    QuotaSelect.FinancingYesNo = other.Financing;
                    QuotaSelect.AmountPaidByCreditCard = other.CreditCard;
                    db.Entry(QuotaSelect).State = EntityState.Modified;
                    db.SaveChanges();                                       
                }
            }
        }));
       

        public QuotationViewModel( ref BuilderContext context, EnumClient res, Quotation select)
        {
            db = context;
            result = res;            
            db.DIC_GroupeQuotations.Load();
            db.DIC_ItemQuotations.Load();
            db.DIC_DescriptionQuotations.Load();            
            db.MaterialQuotations.Load();

            DIC_Section = new List<string>();
            DIC_Section.Add("Material");
            DIC_Section.Add("Labour");
            
            //DIC_Groupes = db.DIC_GroupeQuotations.Local.ToBindingList();
                        
            switch (result)
            {
                case EnumClient.Add:
                {
                        flagCreatQuota = false;
                        QuotaSelect = null;
                }
                break;
                case EnumClient.Ins:
                {
                        flagCreatQuota = true;
                        QuotaSelect = select;
                        ClientSelect = db.Clients.FirstOrDefault(c=>c.Id == QuotaSelect.ClientId);
                        lastIdClient = ClientSelect.Id;
                }
                break;                    
            }            
            
        }
        private void LoadGroup()
        {
            if (DIC_SectionSelect != null)
            {
                if (DIC_SectionSelect == "Material")
                {
                    DIC_Groupes = db.DIC_GroupeQuotations.Local.ToBindingList().Where(g => g.NameGroupe == "FLOORING" || g.NameGroupe == "ACCESSORIES");
                }
                else
                {
                    DIC_Groupes = db.DIC_GroupeQuotations.Local.ToBindingList().Where(g => g.NameGroupe != "FLOORING").Where(g => g.NameGroupe != "ACCESSORIES");
                }
            }
            else
            {
                DIC_Groupes = null;
            }
        }
        private void MaterialView()
        {
            if (MaterialQuotations != null)
            {
                List<MaterialQuotation> labour = new List<MaterialQuotation>();
                List<MaterialQuotation> material = new List<MaterialQuotation>();
                foreach (var item in MaterialQuotations)
                {
                    if (item.Groupe == "FLOORING" || item.Groupe == "ACCESSORIES")
                    {
                        material.Add(new MaterialQuotation()
                        {
                            Id = item.Id,
                            Groupe = item?.Groupe,
                            Item = item?.Item,
                            Description = item?.Description,
                            Quantity = item.Quantity,
                            Price = item.Price,
                            Rate = item.Rate,
                            SupplierId = item?.SupplierId,
                            QuotationId = item.QuotationId,
                            Quotation = item.Quotation
                        });
                    }
                    else
                    {
                        labour.Add(new MaterialQuotation()
                        {
                            Id = item.Id,
                            Groupe = item?.Groupe,
                            Item = item?.Item,
                            Description = item?.Description,
                            Quantity = item.Quantity,
                            Price = item.Price,
                            Rate = item.Rate,
                            QuotationId = item.QuotationId,
                            Quotation = item.Quotation
                        });
                    }
                }
                QuotationsLabour = null;
                QuotationsLabour = labour;
                QuotationsMaterial = null;
                QuotationsMaterial = material;
            }
            else
            {
                QuotationsLabour = null;
                QuotationsMaterial = null;
            }
        }

        
    }
}
