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
using System.Windows.Documents;

namespace Builders.ViewModels
{
    public class QuotationViewModel : ViewModel
    {
        public string WindowName { get; } = "Builder - Quotation";
        public string NameButton { get; } = "  <- Clear    ";
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
        private bool copyClient;    // для копіювання квоти на другого клієнта
        private int? lastIdClient;  // Для збереження попереднього клієнта  

        private decimal quantity;
        private decimal rate;

        private DIC_DepthQuotation depthSelect;
        private IEnumerable<DIC_DepthQuotation> depths;
        private string roomDescription;
        private string companyName;
        private int? mapei;
        private Visibility isVisibleStandart;
        private Visibility isVisibleLeveling;
        private Visibility isVisibleRoomDescription;
       

        private DIC_G_GradeLevel gradeLevelSelect;
        private List<DIC_G_GradeLevel> gradeLevels;
        private DIC_G_Partition partitionSelect;
        private List<DIC_G_Partition> partitions;
        private DIC_G_Additional additionalSelect;
        private List<DIC_G_Additional> additionals;

        private string descriptions;

        private bool isEnableGenerator;
        private bool isEnableCreat;

        private string calcFlooring;
        private string calcAccessories;
        private string calcInstallation;
        private string calcDemolition;
        private string calcServices;

        private GeneratedList generatedListSelect;
        private List<GeneratedList> generatedLists;

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
            get { return materialQuotationSelect; }
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

                    if (CompanyName != "CMO" && MaterialQuotationSelect.Groupe == "FLOORING")
                    {
                        IsVisibleStandart = Visibility.Collapsed;
                        IsVisibleRoomDescription = Visibility.Visible;


                        DIC_GroupeSelect = (MaterialQuotationSelect?.Groupe != null) ? DIC_Groupes?.FirstOrDefault(g => g?.NameGroupe == MaterialQuotationSelect?.Groupe) : null;
                        RoomDescription = MaterialQuotationSelect?.Description;
                        DepthSelect = (MaterialQuotationSelect?.Depth != null) ? Depths?.FirstOrDefault(g => g?.Name == MaterialQuotationSelect?.Depth) : null;
                        Quantity = MaterialQuotationSelect?.QuantityNL ?? 0m;
                        Rate = MaterialQuotationSelect?.Rate ?? 0m;
                    }
                    else
                    {
                        IsVisibleStandart = Visibility.Visible;
                        IsVisibleRoomDescription = Visibility.Collapsed;
                        
                        DIC_GroupeSelect = (MaterialQuotationSelect?.Groupe != null) ? DIC_Groupes?.FirstOrDefault(g => g?.NameGroupe == MaterialQuotationSelect?.Groupe) : null;
                        DIC_ItemSelect = (MaterialQuotationSelect?.Item != null) ? DIC_Items?.FirstOrDefault(i => i?.Name == MaterialQuotationSelect?.Item) : null;
                        DIC_DescriptionSelect = (MaterialQuotationSelect?.Description != null) ? DIC_Descriptions?.FirstOrDefault(d => d?.Name == MaterialQuotationSelect?.Description) : null;
                        //Quantity = MaterialQuotationSelect?.Quantity ?? 0m;
                        //Rate = MaterialQuotationSelect?.Rate ?? 0m;
                    }
                }
            }
        }
        public IEnumerable<MaterialQuotation> MaterialQuotations
        {
            get { return materialQuotations; }
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
                if (CompanyName != "CMO" && DIC_GroupeSelect != null && DIC_GroupeSelect.NameGroupe == "FLOORING")
                {
                    IsVisibleStandart = Visibility.Collapsed;
                    IsVisibleRoomDescription = Visibility.Visible;
                    //DIC_Items = null;
                }
                else
                {
                    IsVisibleStandart = Visibility.Visible;
                    IsVisibleRoomDescription = Visibility.Collapsed;
                    DIC_Items = (DIC_GroupeSelect != null) ? db.DIC_ItemQuotations.Local.ToBindingList().Where(i => i.GroupeId == DIC_GroupeSelect.Id).OrderBy(i => i.Name) : null;
                }
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

        public DIC_DepthQuotation DepthSelect
        {
            get { return depthSelect; }
            set
            {
                depthSelect = value;
                OnPropertyChanged(nameof(DepthSelect));
            }
        }
        public IEnumerable<DIC_DepthQuotation> Depths
        {
            get { return depths; }
            set
            {
                depths = value;
                OnPropertyChanged(nameof(Depths));
            }
        }
        public string RoomDescription
        {
            get { return roomDescription; }
            set
            {
                roomDescription = value;
                OnPropertyChanged(nameof(RoomDescription));
            }
        }
        public string CompanyName
        {
            get { return companyName; }
            set
            {
                companyName = value;
                OnPropertyChanged(nameof(CompanyName));
               
            }
        }
        public int? Mapei
        {
            get { return mapei; }
            set
            {
                mapei = value;
                OnPropertyChanged(nameof(Mapei));
            }
        }
        public Visibility IsVisibleStandart
        {
            get { return isVisibleStandart; }
            set
            {
                isVisibleStandart = value;
                OnPropertyChanged(nameof(IsVisibleStandart));
            }
        }
        public Visibility IsVisibleLeveling
        {
            get { return isVisibleLeveling; }
            set
            {
                isVisibleLeveling = value;
                OnPropertyChanged(nameof(IsVisibleLeveling));
            }
        }
        public Visibility IsVisibleRoomDescription
        {
            get { return isVisibleRoomDescription; }
            set
            {
                isVisibleRoomDescription = value;
                OnPropertyChanged(nameof(IsVisibleRoomDescription));
            }
        }
        

        public DIC_G_GradeLevel GradeLevelSelect
        {
            get { return gradeLevelSelect; }
            set
            {
                gradeLevelSelect = value;
                OnPropertyChanged(nameof(GradeLevelSelect));
            }
        }
        public List<DIC_G_GradeLevel> GradeLevels
        {
            get { return gradeLevels; }
            set
            {
                gradeLevels = value;
                OnPropertyChanged(nameof(GradeLevels));
            }
        }
        public DIC_G_Partition PartitionSelect
        {
            get { return partitionSelect; }
            set
            {
                partitionSelect = value;
                OnPropertyChanged(nameof(PartitionSelect));
            }
        }
        public List<DIC_G_Partition> Partitions
        {
            get { return partitions; }
            set
            {
                partitions = value;
                OnPropertyChanged(nameof(Partitions));
            }
        }
        public DIC_G_Additional AdditionalSelect
        {
            get { return additionalSelect; }
            set
            {
                additionalSelect = value;
                OnPropertyChanged(nameof(AdditionalSelect));
            }
        }
        public List<DIC_G_Additional> Additionals
        {
            get { return additionals; }
            set
            {
                additionals = value;
                OnPropertyChanged(nameof(Additionals));
            }
        }

        public string Descriptions
        {
            get { return descriptions; }
            set
            {
                descriptions = value;
                OnPropertyChanged(nameof(Descriptions));
            }
        }

        public bool IsEnableGenerator
        {
            get { return isEnableGenerator; }
            set
            {
                isEnableGenerator = value;
                OnPropertyChanged(nameof(IsEnableGenerator));
                if (IsEnableGenerator)
                {
                    IsEnableCreat = false;
                }
                else
                {
                    IsEnableCreat = true;
                }
            }
        }
        public bool IsEnableCreat
        {
            get { return isEnableCreat; }
            set
            {
                isEnableCreat = value;
                OnPropertyChanged(nameof(IsEnableCreat));
            }
        }

        public string CalcFlooring
        {
            get { return calcFlooring; }
            set
            {
                calcFlooring = value;
                OnPropertyChanged(nameof(CalcFlooring));
            }
        }
        public string CalcAccessories
        {
            get { return calcAccessories; }
            set
            {
                calcAccessories = value;
                OnPropertyChanged(nameof(CalcAccessories));
            }
        }
        public string CalcInstallation
        {
            get { return calcInstallation; }
            set
            {
                calcInstallation = value;
                OnPropertyChanged(nameof(CalcInstallation));
            }
        }
        public string CalcDemolition
        {
            get { return calcDemolition; }
            set
            {
                calcDemolition = value;
                OnPropertyChanged(nameof(CalcDemolition));
            }
        }
        public string CalcServices
        {
            get { return calcServices; }
            set
            {
                calcServices = value;
                OnPropertyChanged(nameof(CalcServices));
            }
        }

        public GeneratedList GeneratedListSelect
        {
            get { return generatedListSelect; }
            set
            {
                generatedListSelect = value;
                OnPropertyChanged(nameof(GeneratedListSelect));
            }
        }
        public List<GeneratedList> GeneratedLists
        {
            get { return generatedLists; }
            set
            {
                generatedLists = value;
                OnPropertyChanged(nameof(GeneratedLists));
            }
        }
        //**************************************************************************************************
        private Command _addItem;
        private Command _insItem;
        private Command _delItem;
        private Command _otherQuotation;
        private Command _clearRoom;
        private Command _generated;
        private Command _search;



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
                    JobDescription = Descriptions,
                    JobNote = "",

                    CompanyName = CompanyName,

                    MaterialSubtotal = 0m,
                    MaterialDiscountYN = 0m,
                    MaterialDiscountAmount = 0m,

                    LabourSubtotal = 0m,
                    LabourDiscountYN = 0m,
                    LabourDiscountAmount = 0m,

                    FinancingYesNo = false,
                    AmountPaidByCreditCard = 0m

                };
                db.Quotations.Add(quotation);
                db.SaveChanges();
                flagCreatQuota = true;
                QuotaSelect = db.Quotations.Local.ToBindingList().OrderByDescending(q => q.Id).FirstOrDefault();
                QuotaId = QuotaSelect.Id;
                IsEnableGenerator = true;

                Rate = 0m;
                Quantity = 0m;
            }
            if (flagCreatQuota && QuotaSelect != null)
            {

                var temp = MaterialQuotations; //db.MaterialQuotations.Where(f => f.QuotationId == QuotaSelect.Id);

                int flooring = (temp != null) ? (temp.Where(t => t.Groupe == "FLOORING").ToList().Select(t=>t.MaterialDetail).Distinct().Count()) : 0;
                int accessories = (temp != null) ? (temp.Where(t => t.Groupe == "ACCESSORIES").ToList().Select(t => t.MaterialDetail).Distinct().Count()) : 0;
                int installation = (temp != null) ? (temp.Where(t => t.Groupe == "INSTALLATION").ToList().Select(t => t.MaterialDetail).Distinct().Count()) : 0;
                int demolition = (temp != null) ? (temp.Where(t => t.Groupe == "DEMOLITION").ToList().Select(t => t.MaterialDetail).Distinct().Count()) : 0;
                int services = (temp != null) ? (temp.Where(t => t.Groupe == "OPTIONAL SERVICES").ToList().Select(t => t.MaterialDetail).Distinct().Count()) : 0;
                int delivery = (temp != null) ? (temp.Where(t => t.Groupe == "FLOORING DELIVERY").ToList().Select(t => t.MaterialDetail).Distinct().Count()) : 0;

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
                    if (CompanyName != "CMO" && DIC_GroupeSelect.NameGroupe == "FLOORING")
                    {
                        int mapei = 0;
                        if (Quantity > 0 && DepthSelect != null)
                        {
                            mapei = (int)decimal.Ceiling(Quantity / DepthSelect.Price);
                        }
                        else
                        {
                            mapei = 0;
                        }

                        MaterialQuotation material = new MaterialQuotation()
                        {
                            QuotationId = QuotaSelect.Id,
                            Groupe = DIC_GroupeSelect?.NameGroupe,
                            Item = null,
                            Description = RoomDescription,
                            Mapei = mapei,
                            Depth = DepthSelect?.Name,
                            Rate = decimal.Round(Rate, 2),
                            Quantity = mapei,
                            QuantityNL = Quantity,
                            Price = decimal.Round(Rate * mapei, 2),
                            SupplierId = DIC_ItemSelect?.SupplierId,
                            GradeLevel = GradeLevelSelect?.Name,
                            Partition = PartitionSelect?.Name,
                            Aditional = AdditionalSelect?.Name
                        };
                        db.MaterialQuotations.Add(material);
                        db.SaveChanges();

                        MaterialQuotations = null;
                        MaterialQuotations = db.MaterialQuotations.Local.ToBindingList().Where(m => m.QuotationId == QuotaSelect.Id);
                        MaterialQuotationSelect = material;
                    }
                    else
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
                            SupplierId = DIC_ItemSelect?.SupplierId,
                            GradeLevel = GradeLevelSelect?.Name,
                            Partition = PartitionSelect?.Name,
                            Aditional = AdditionalSelect?.Name
                        };
                        db.MaterialQuotations.Add(material);
                        db.SaveChanges();

                        MaterialQuotations = null;
                        MaterialQuotations = db.MaterialQuotations.Local.ToBindingList().Where(m => m.QuotationId == QuotaSelect.Id);
                        MaterialQuotationSelect = material;
                        Rate = 0m;
                        Quantity = 0m;
                    }
                }
            }
        }));
        public Command InsItem => _insItem ?? (_insItem = new Command(obj =>
        {
            if (MaterialQuotationSelect != null)
            {
                var temp = MaterialQuotations; //db.MaterialQuotations.Where(f => f.QuotationId == QuotaSelect.Id);
                var tempMaterialSelect = db.MaterialQuotations.FirstOrDefault(m => m.Id == MaterialQuotationSelect.Id);

                int flooring = (temp != null) ? (temp.Where(t => t.Groupe == "FLOORING").ToList().Select(t => t.MaterialDetail).Distinct().Count()) : 0;
                int accessories = (temp != null) ? (temp.Where(t => t.Groupe == "ACCESSORIES").ToList().Select(t => t.MaterialDetail).Distinct().Count()) : 0;
                int installation = (temp != null) ? (temp.Where(t => t.Groupe == "INSTALLATION").ToList().Select(t => t.MaterialDetail).Distinct().Count()) : 0;
                int demolition = (temp != null) ? (temp.Where(t => t.Groupe == "DEMOLITION").ToList().Select(t => t.MaterialDetail).Distinct().Count()) : 0;
                int services = (temp != null) ? (temp.Where(t => t.Groupe == "OPTIONAL SERVICES").ToList().Select(t => t.MaterialDetail).Distinct().Count()) : 0;
                int delivery = (temp != null) ? (temp.Where(t => t.Groupe == "FLOORING DELIVERY").ToList().Select(t => t.MaterialDetail).Distinct().Count()) : 0;

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
                    if (CompanyName != "CMO" && DIC_GroupeSelect.NameGroupe == "FLOORING")
                    {
                        int mapei = 0;
                        if (Quantity > 0 && DepthSelect != null)
                        {
                            mapei = (int)decimal.Ceiling(Quantity / DepthSelect.Price);
                        }
                        else
                        {
                            mapei = 0;
                        }

                        tempMaterialSelect.Groupe = DIC_GroupeSelect?.NameGroupe;
                        tempMaterialSelect.Item = null;
                        tempMaterialSelect.SupplierId = DIC_ItemSelect?.SupplierId;
                        tempMaterialSelect.Description = RoomDescription;
                        tempMaterialSelect.Mapei = mapei;
                        tempMaterialSelect.Depth = DepthSelect?.Name;
                        tempMaterialSelect.Rate = decimal.Round(Rate, 2);
                        tempMaterialSelect.Quantity = mapei;
                        tempMaterialSelect.QuantityNL = Quantity;
                        tempMaterialSelect.Price = decimal.Round(Rate * mapei, 2);
                        tempMaterialSelect.GradeLevel = GradeLevelSelect?.Name;
                        tempMaterialSelect.Partition = PartitionSelect?.Name;
                        tempMaterialSelect.Aditional = AdditionalSelect?.Name;

                        db.Entry(tempMaterialSelect).State = EntityState.Modified;
                        db.SaveChanges();
                        MaterialQuotationSelect = null;
                        MaterialQuotations = (QuotaSelect != null) ? (db.MaterialQuotations.Local.ToBindingList().Where(m => m.QuotationId == QuotaSelect.Id)) : null;

                    }
                    else
                    {
                        tempMaterialSelect.Groupe = DIC_GroupeSelect?.NameGroupe;
                        tempMaterialSelect.Item = DIC_ItemSelect?.Name;
                        tempMaterialSelect.SupplierId = DIC_ItemSelect?.SupplierId;
                        tempMaterialSelect.Description = DIC_DescriptionSelect?.Name;
                        tempMaterialSelect.Rate = decimal.Round(Rate, 2);
                        tempMaterialSelect.Quantity = Quantity;
                        tempMaterialSelect.Price = decimal.Round(Rate * Quantity, 2);
                        tempMaterialSelect.GradeLevel = GradeLevelSelect?.Name;
                        tempMaterialSelect.Partition = PartitionSelect?.Name;
                        tempMaterialSelect.Aditional = AdditionalSelect?.Name;

                        db.Entry(tempMaterialSelect).State = EntityState.Modified;
                        db.SaveChanges();
                        MaterialQuotationSelect = null;
                        MaterialQuotations = (QuotaSelect != null) ? (db.MaterialQuotations.Local.ToBindingList().Where(m => m.QuotationId == QuotaSelect.Id)) : null;
                        Rate = 0m;
                        Quantity = 0m;
                    }
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
                    Descriptions = other.Description;
                }
            }
        }));
        public Command ClearRoom => _clearRoom ?? (_clearRoom = new Command(obj=> 
        {
            GradeLevelSelect = null;
            PartitionSelect = null;
            AdditionalSelect = null;
        }));
        public Command Generated => _generated ?? (_generated = new Command(async obj=> 
        {
            GeneratedLists = null;
            var displayRootRegistry = (Application.Current as App).displayRootRegistry;
            var generator = new GeneratedViewModel(ref db, QuotaSelect.Id);
            await displayRootRegistry.ShowModalPresentation(generator);
            GeneratedLists = CalculateGenerated(generator.generatedSelect);
            QuotaSelect = AddMaterialQuotaToGenerated(GeneratedLists, quotaSelect, generator.OptionQuota);
        }));
        public Command Search => _search ?? (_search = new Command(obj=> 
        {
            string search = obj.ToString();
            if (search == "")
            {
                DIC_Items = (DIC_GroupeSelect != null) ? db.DIC_ItemQuotations.Local.ToBindingList().Where(i => i.GroupeId == DIC_GroupeSelect.Id).OrderBy(i => i.Name) : null;
                DIC_ItemSelect = null;
            }
            else
            {
                DIC_Items = (DIC_GroupeSelect != null) ? (DIC_Items.Where(t => t.GroupeId == DIC_GroupeSelect.Id)
                                                                   .Where(t => t.Name.ToUpper().Contains(search.ToUpper()))
                                                                   .OrderBy(t => t.Name)) : null;
            }
        }));

       
        public QuotationViewModel(ref BuilderContext context, EnumClient res, Quotation select, string companyName)
        {
            db = context;
            result = res;
            CompanyName = companyName;
            IsEnableGenerator = false;
            db.DIC_GroupeQuotations.Load();
            db.DIC_ItemQuotations.Load();
            db.DIC_DescriptionQuotations.Load();
            db.DIC_DepthQuotations.Load();
            db.MaterialQuotations.Load();

            Depths = db.DIC_DepthQuotations.Local.ToBindingList();

            DIC_Section = new List<string>();
            DIC_Section.Add("Material");
            DIC_Section.Add("Labour");

            LoadComboBox();

            IsVisibleStandart = Visibility.Visible;
            IsVisibleRoomDescription = Visibility.Collapsed;
                       
            
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
                        IsEnableGenerator = true;
                        flagCreatQuota = true;
                        QuotaSelect = select;
                        Descriptions = select?.JobDescription;
                        ClientSelect = db.Clients.FirstOrDefault(c => c.Id == QuotaSelect.ClientId);
                        lastIdClient = ClientSelect.Id;
                        var generator = db.Generateds.FirstOrDefault(g => g.QuotaId == select.Id);
                        if (generator != null)
                        {
                            GeneratedLists = db.GeneratedLists.Where(g => g.GeneratedId == generator.Id).ToList();
                        }
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
        private void LoadComboBox()
        {
            GradeLevels = new List<DIC_G_GradeLevel>();
            Partitions = new List<DIC_G_Partition>();
            Additionals = new List<DIC_G_Additional>();

            GradeLevels = db.DIC_G_GradeLevels.ToList();
            Partitions = db.DIC_G_Partitions.ToList();
            Additionals = db.DIC_G_Additionals.ToList();
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
                            Depth = item?.Depth,
                            Mapei = item?.Mapei,
                            Quantity = item.Quantity,
                            QuantityNL = item.QuantityNL,
                            Price = item.Price,
                            Rate = item.Rate,
                            SupplierId = item?.SupplierId,
                            QuotationId = item.QuotationId,
                            Quotation = item.Quotation,
                            GradeLevel = item?.GradeLevel,
                            Partition = item?.Partition,
                            Aditional = item?.Aditional
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
                            Depth = item?.Depth,
                            Mapei = item?.Mapei,
                            Quantity = item.Quantity,
                            QuantityNL = item.QuantityNL,
                            Price = item.Price,
                            Rate = item.Rate,
                            QuotationId = item.QuotationId,
                            Quotation = item.Quotation,
                            GradeLevel = item?.GradeLevel,
                            Partition = item?.Partition,
                            Aditional = item?.Aditional
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
        private List<GeneratedList> CalculateGenerated(Generated select)
        {
            var temp = db.GeneratedLists.Where(g => g.GeneratedId == select.Id);
            if (temp != null)
            {
                db.GeneratedLists.RemoveRange(temp);
                db.SaveChanges();
            }


            var material = db.GeneratedMaterials.Where(m => m.GeneratedId == select.Id).ToList();
            var accessories = db.GeneratedAccessories.Where(a => a.GeneratedId == select.Id).ToList();
            var stairs = db.GeneratedStairs.Where(s => s.GeneratedId == select.Id).ToList();
            var molding = db.GeneratedMoldings.Where(m => m.GeneratedId == select.Id).ToList();
            var suplementary = db.GeneratedSuplementaries.Where(s => s.GeneratedId == select.Id).ToList();
            var flood = db.GeneratedFloods.Where(f => f.GeneratedId == select.Id).ToList();

            GeneratedList list = new GeneratedList();
            List<GeneratedList> result = new List<GeneratedList>();

            foreach (var item in material)
            {
                if (item.NewFloor != null)
                {
                    // Перевіряєм чи вже є такий запис, якщо є то додаєм тільки кількість. Не робим дублів. 
                    var totalResult = result.FirstOrDefault(res => res.GeneratedId == select.Id &&
                                                                   res.Groupe == "FLOORING" &&
                                                                   res.GradeLevel == item?.GradeLevel &&
                                                                   res.Partition == item?.Partition &&
                                                                   res.Aditional == item?.Aditional &&
                                                                   res.Name == item?.NewFloor);
                    if (totalResult != null)
                    {
                        int index = result.IndexOf(totalResult);                        
                        if (index != -1)
                        {
                            result[index].Count += item.TotalFlooringMaterial;
                        }
                        
                    }
                    else
                    {
                        list.GeneratedId = select.Id;
                        list.Groupe = "FLOORING";
                        list.GradeLevel = item?.GradeLevel;
                        list.Partition = item?.Partition;
                        list.Aditional = item?.Aditional;
                        list.Name = item?.NewFloor;
                        list.Count = item.TotalFlooringMaterial;

                        result.Add(list);
                    }
                    list = null;
                    list = new GeneratedList();

                    // Перевіряєм чи вже є такий запис, якщо є то додаєм тільки кількість. Не робим дублів. 
                    totalResult = result.FirstOrDefault(res => res.GeneratedId == select.Id &&
                                                                   res.Groupe == "INSTALLATION" &&
                                                                   res.GradeLevel == item?.GradeLevel &&
                                                                   res.Partition == item?.Partition &&
                                                                   res.Aditional == item?.Aditional &&
                                                                   res.Name == item?.NewFloor);
                    if (totalResult != null)
                    {
                        totalResult.Count += item.TotalFloor;
                        result.RemoveAt(result.IndexOf(totalResult));
                    }
                    else
                    {

                        list.GeneratedId = select.Id;
                        list.Groupe = "INSTALLATION";
                        list.GradeLevel = item?.GradeLevel;
                        list.Partition = item?.Partition;
                        list.Aditional = item?.Aditional;
                        list.Name = item?.NewFloor;
                        list.Count = item.TotalFloor;

                        result.Add(list);
                    }
                    list = null;
                    list = new GeneratedList();
                }

                if (item.NeedDemolition == "Yes")
                {
                    // Перевіряєм чи вже є такий запис, якщо є то додаєм тільки кількість. Не робим дублів. 
                    var totalResult = result.FirstOrDefault(res => res.GeneratedId == select.Id &&
                                                                   res.Groupe == "DEMOLITION" &&
                                                                   res.GradeLevel == item?.GradeLevel &&
                                                                   res.Partition == item?.Partition &&
                                                                   res.Aditional == item?.Aditional &&
                                                                   res.Name == item?.ExistingFloor);
                    if (totalResult != null)
                    {
                        int index = result.IndexOf(totalResult);
                        if (index != -1)
                        {
                            result[index].Count += item.TotalFloor;
                        }
                    }
                    else
                    {
                        list.GeneratedId = select.Id;
                        list.Groupe = "DEMOLITION";
                        list.GradeLevel = item?.GradeLevel;
                        list.Partition = item?.Partition;
                        list.Aditional = item?.Aditional;
                        list.Name = item?.ExistingFloor;
                        list.Count = item.TotalFloor;

                        result.Add(list);
                    }
                    list = null;
                    list = new GeneratedList();
                }

                if (item.TransitionR != 0m)
                {
                    // Перевіряєм чи вже є такий запис, якщо є то додаєм тільки кількість. Не робим дублів. 
                    var totalResult = result.FirstOrDefault(res => res.GeneratedId == select.Id &&
                                                                   res.Groupe == "ACCESSORIES" &&
                                                                   res.GradeLevel == item?.GradeLevel &&
                                                                   res.Partition == item?.Partition &&
                                                                   res.Aditional == item?.Aditional &&
                                                                   res.Name == "Transition (R)");
                    if (totalResult != null)
                    {
                        int index = result.IndexOf(totalResult);
                        if (index != -1)
                        {
                            result[index].Count += item.TransitionR;
                        }
                    }
                    else
                    {
                        list.GeneratedId = select.Id;
                        list.Groupe = "ACCESSORIES";
                        list.GradeLevel = item?.GradeLevel;
                        list.Partition = item?.Partition;
                        list.Aditional = item?.Aditional;
                        list.Name = "Transition (R)";
                        list.Count = item.TransitionR;

                        result.Add(list);
                    }
                    list = null;
                    list = new GeneratedList();
                }

                if (item.TransitionT != 0m)
                {
                    // Перевіряєм чи вже є такий запис, якщо є то додаєм тільки кількість. Не робим дублів. 
                    var totalResult = result.FirstOrDefault(res => res.GeneratedId == select.Id &&
                                                                   res.Groupe == "ACCESSORIES" &&
                                                                   res.GradeLevel == item?.GradeLevel &&
                                                                   res.Partition == item?.Partition &&
                                                                   res.Aditional == item?.Aditional &&
                                                                   res.Name == "Transition (T)");
                    if (totalResult != null)
                    {
                        int index = result.IndexOf(totalResult);
                        if (index != -1)
                        {
                            result[index].Count += item.TransitionT;
                        }
                    }
                    else
                    {
                        list.GeneratedId = select.Id;
                        list.Groupe = "ACCESSORIES";
                        list.GradeLevel = item?.GradeLevel;
                        list.Partition = item?.Partition;
                        list.Aditional = item?.Aditional;
                        list.Name = "Transition (T)";
                        list.Count = item.TransitionT;

                        result.Add(list);
                    }
                    list = null;
                    list = new GeneratedList();
                }

                if (item.TransitionOther != 0m)
                {
                    // Перевіряєм чи вже є такий запис, якщо є то додаєм тільки кількість. Не робим дублів. 
                    var totalResult = result.FirstOrDefault(res => res.GeneratedId == select.Id &&
                                                                   res.Groupe == "ACCESSORIES" &&
                                                                   res.GradeLevel == item?.GradeLevel &&
                                                                   res.Partition == item?.Partition &&
                                                                   res.Aditional == item?.Aditional &&
                                                                   res.Name == item.NoteTransitionOther);
                    if (totalResult != null)
                    {
                        int index = result.IndexOf(totalResult);
                        if (index != -1)
                        {
                            result[index].Count += item.TransitionOther;
                        }
                    }
                    else
                    {
                        list.GeneratedId = select.Id;
                        list.Groupe = "ACCESSORIES";
                        list.GradeLevel = item?.GradeLevel;
                        list.Partition = item?.Partition;
                        list.Aditional = item?.Aditional;
                        list.Name = item.NoteTransitionOther;
                        list.Count = item.TransitionOther;

                        result.Add(list);
                    }
                    list = null;
                    list = new GeneratedList();
                }

                if (item.QtyTrim != 0m)
                {
                    // Перевіряєм чи вже є такий запис, якщо є то додаєм тільки кількість. Не робим дублів. 
                    var totalResult = result.FirstOrDefault(res => res.GeneratedId == select.Id &&
                                                                   res.Groupe == "ACCESSORIES" &&
                                                                   res.GradeLevel == item?.GradeLevel &&
                                                                   res.Partition == item?.Partition &&
                                                                   res.Aditional == item?.Aditional &&
                                                                   res.Name == item.TypeTrim);
                    if (totalResult != null)
                    {
                        int index = result.IndexOf(totalResult);
                        if (index != -1)
                        {
                            result[index].Count += item.QtyTrim;
                        }
                    }
                    else
                    {
                        list.GeneratedId = select.Id;
                        list.Groupe = "ACCESSORIES";
                        list.GradeLevel = item?.GradeLevel;
                        list.Partition = item?.Partition;
                        list.Aditional = item?.Aditional;
                        list.Name = item.TypeTrim;
                        list.Count = item.QtyTrim;

                        result.Add(list);
                    }
                    list = null;
                    list = new GeneratedList();

                    // Перевіряєм чи вже є такий запис, якщо є то додаєм тільки кількість. Не робим дублів. 
                    totalResult = result.FirstOrDefault(res => res.GeneratedId == select.Id &&
                                                                res.Groupe == "INSTALLATION" &&
                                                                res.GradeLevel == item?.GradeLevel &&
                                                                res.Partition == item?.Partition &&
                                                                res.Aditional == item?.Aditional &&
                                                                res.Name == item.TypeTrim);
                    if (totalResult != null)
                    {
                        totalResult.Count += item.QtyTrim;
                        result.RemoveAt(result.IndexOf(totalResult));
                    }
                    else
                    {
                        list.GeneratedId = select.Id;
                        list.Groupe = "INSTALLATION";
                        list.GradeLevel = item?.GradeLevel;
                        list.Partition = item?.Partition;
                        list.Aditional = item?.Aditional;
                        list.Name = item.TypeTrim;
                        list.Count = item.QtyTrim;

                        result.Add(list);
                    }
                    list = null;
                    list = new GeneratedList();
                }
            }

            decimal r = material.Select(m => m.TransitionR).Sum();
            decimal t = material.Select(m => m.TransitionT).Sum();
            decimal o = material.Select(m => m.TransitionOther).Sum();

            if ((r + t + o) != 0m)
            {
                list.GeneratedId = select.Id;
                list.Groupe = "INSTALLATION";
                list.Name = "Total transition";
                list.Count = r + t + o;

                result.Add(list);
                list = null;
                list = new GeneratedList();
            }

            foreach (var item in accessories)
            {
                if (item.AccessoriesFloor != null)
                {
                    // Перевіряєм чи вже є такий запис, якщо є то додаєм тільки кількість. Не робим дублів. 
                    var totalResult = result.FirstOrDefault(res => res.GeneratedId == select.Id &&
                                                                   res.Groupe == "ACCESSORIES" &&                                                                   
                                                                   res.Name == item.AccessoriesFloor + "  " + item.NewFloor + "-" + item.TypeAccessoriesFloor);
                    if (totalResult != null)
                    {
                        int index = result.IndexOf(totalResult);
                        if (index != -1)
                        {
                            result[index].Count += item.TotalAccessoriesFloor;
                        }
                    }
                    else
                    {
                        list.GeneratedId = select.Id;
                        list.Groupe = "ACCESSORIES";
                        list.Name = item.AccessoriesFloor + "  " + item.NewFloor + "-" + item.TypeAccessoriesFloor;
                        list.Count = item.TotalAccessoriesFloor;

                        result.Add(list);
                    }
                    list = null;
                    list = new GeneratedList();
                }
            }

            foreach (var item in stairs)
            {
                if (item.TypeStairs != null)
                {
                    // Перевіряєм чи вже є такий запис, якщо є то додаєм тільки кількість. Не робим дублів. 
                    var totalResult = result.FirstOrDefault(res => res.GeneratedId == select.Id &&
                                                                   res.Groupe == "ACCESSORIES" &&
                                                                   res.GradeLevel == item?.GradeLevel &&                                                                   
                                                                   res.Name == item.TypeStairs + "  LENGHT (FT) - " + item.LenghtStairs.ToString());
                    if (totalResult != null)
                    {
                        int index = result.IndexOf(totalResult);
                        if (index != -1)
                        {
                            result[index].Count += item.QtyStairsLenght;
                        }
                    }
                    else
                    {
                        list.GeneratedId = select.Id;
                        list.GradeLevel = item.GradeLevel;
                        list.Groupe = "ACCESSORIES";
                        list.Name = item.TypeStairs + "  LENGHT (FT) - " + item.LenghtStairs.ToString();
                        list.Count = item.QtyStairsLenght;

                        result.Add(list);
                    }
                    list = null;
                    list = new GeneratedList();

                    // Перевіряєм чи вже є такий запис, якщо є то додаєм тільки кількість. Не робим дублів. 
                    totalResult = result.FirstOrDefault(res => res.GeneratedId == select.Id &&
                                                                   res.Groupe == "INSTALLATION" &&
                                                                   res.GradeLevel == item?.GradeLevel &&
                                                                   res.Name == item.TypeStairs + "  LENGHT (FT) - " + item.LenghtStairs.ToString());
                    if (totalResult != null)
                    {
                        int index = result.IndexOf(totalResult);
                        if (index != -1)
                        {
                            result[index].Count += item.QtyStairs;
                        }
                    }
                    else
                    {
                        list.GeneratedId = select.Id;
                        list.GradeLevel = item.GradeLevel;
                        list.Groupe = "INSTALLATION";
                        list.Name = item.TypeStairs + "  LENGHT (FT) - " + item.LenghtStairs.ToString();
                        list.Count = item.QtyStairs;

                        result.Add(list);
                    }
                    list = null;
                    list = new GeneratedList();
                }

                if (item.TypeLeveling != null)
                {
                    // Перевіряєм чи вже є такий запис, якщо є то додаєм тільки кількість. Не робим дублів. 
                    var totalResult = result.FirstOrDefault(res => res.GeneratedId == select.Id &&
                                                                   res.Groupe == "ACCESSORIES" &&
                                                                   res.GradeLevel == item?.GradeLevel &&
                                                                   res.Name == item.NameLeveling);
                    if (totalResult != null)
                    {
                        int index = result.IndexOf(totalResult);
                        if (index != -1)
                        {
                            result[index].Count += item.QtyLeveling;
                        }
                    }
                    else
                    {
                        list.GeneratedId = select.Id;
                        list.GradeLevel = item.GradeLevel;
                        list.Groupe = "ACCESSORIES";
                        list.Name = item.NameLeveling;
                        list.Count = item.QtyLeveling;

                        result.Add(list);
                    }
                    list = null;
                    list = new GeneratedList();
                }
            }

            decimal sumQtyStair = stairs.Where(s=>s.Demolition == "Yes").ToList()?.Select(s => s.QtyStairs)?.Sum() ?? 0m;
            if (sumQtyStair != 0m)
            {
                list.GeneratedId = select.Id;
                list.Groupe = "DEMOLITION";
                list.Name = "Stairs";
                list.Count = sumQtyStair;

                result.Add(list);
                list = null;
                list = new GeneratedList();
            }

            GeneratedList listPainting = new GeneratedList();
            foreach (var item in molding)
            {
                if (item.MoldingName != null)
                {
                    // Перевіряєм чи вже є такий запис, якщо є то додаєм тільки кількість. Не робим дублів. 
                    var totalResult = result.FirstOrDefault(res => res.GeneratedId == select.Id &&
                                                                   res.Groupe == "ACCESSORIES" &&                                                                   
                                                                   res.Name == item.MoldingName + "  -  " + item.TypeMolding);
                    if (totalResult != null)
                    {
                        int index = result.IndexOf(totalResult);
                        if (index != -1)
                        {
                            result[index].Count += item.BaseboardMaterial;
                        }
                    }
                    else
                    {
                        list.GeneratedId = select.Id;
                        list.Groupe = "ACCESSORIES";
                        list.Name = item.MoldingName + "  -  " + item.TypeMolding;
                        list.Count = item.BaseboardMaterial;

                        result.Add(list);
                    }
                    list = null;
                    list = new GeneratedList();

                    // Перевіряєм чи вже є такий запис, якщо є то додаєм тільки кількість. Не робим дублів. 
                    totalResult = result.FirstOrDefault(res => res.GeneratedId == select.Id &&
                                                                   res.Groupe == "INSTALLATION" &&
                                                                   res.Name == item.MoldingName + "  -  " + item.TypeMolding);
                    if (totalResult != null)
                    {
                        int index = result.IndexOf(totalResult);
                        if (index != -1)
                        {
                            result[index].Count += item.BaseboardLabour;
                        }
                    }
                    else
                    {
                        list.GeneratedId = select.Id;
                        list.Groupe = "INSTALLATION";
                        list.Name = item.MoldingName + "  -  " + item.TypeMolding;
                        list.Count = item.BaseboardLabour;

                        result.Add(list);
                    }
                    list = null;
                    list = new GeneratedList();
                }
                if (item.Painting != null)
                {
                    var name = result.FirstOrDefault(l => l.Groupe == "OPTIONAL SERVICES" && l.Name == item.Painting);
                    if (name == null)
                    {
                        listPainting.GeneratedId = select.Id;
                        listPainting.Groupe = "OPTIONAL SERVICES";
                        listPainting.Name = item.Painting;
                        listPainting.Count = item.BaseboardMaterial;

                        result.Add(listPainting);
                        listPainting = null;
                        listPainting = new GeneratedList();
                    }
                    else
                    {
                        listPainting.GeneratedId = select.Id;
                        listPainting.Groupe = "OPTIONAL SERVICES";
                        listPainting.Name = name.Name;
                        listPainting.Count = item.BaseboardMaterial + name.Count;

                        result.RemoveAt(result.IndexOf(name));     // Видаляє елемент з колекції
                        result.Add(listPainting);
                        listPainting = null;
                        listPainting = new GeneratedList();
                    }                    
                }
            }

            foreach (var item in suplementary)
            {
                if (item.FurnitureHandelingRoom != 0m)
                {
                    // Перевіряєм чи вже є такий запис, якщо є то додаєм тільки кількість. Не робим дублів. 
                    var totalResult = result.FirstOrDefault(res => res.GeneratedId == select.Id &&
                                                                   res.Groupe == "OPTIONAL SERVICES" &&
                                                                   res.Name == "Furniture Moving");
                    if (totalResult != null)
                    {
                        int index = result.IndexOf(totalResult);
                        if (index != -1)
                        {
                            result[index].Count += item.FurnitureHandelingRoom;
                        }
                    }
                    else
                    {
                        list.GeneratedId = select.Id;
                        list.Groupe = "OPTIONAL SERVICES";
                        list.Name = "Furniture Moving";
                        list.Count = item.FurnitureHandelingRoom;

                        result.Add(list);
                    }
                    list = null;
                    list = new GeneratedList();
                }
                if (item.RateDisposal != 0m)
                {
                    // Перевіряєм чи вже є такий запис, якщо є то додаєм тільки кількість. Не робим дублів. 
                    var totalResult = result.FirstOrDefault(res => res.GeneratedId == select.Id &&
                                                                   res.Groupe == "DEMOLITION" &&
                                                                   res.Name == "Disposal");
                    if (totalResult != null)
                    {
                        int index = result.IndexOf(totalResult);
                        if (index != -1)
                        {
                            result[index].Count += item.RateDisposal;
                        }
                    }
                    else
                    {
                        list.GeneratedId = select.Id;
                        list.Groupe = "DEMOLITION";
                        list.Name = "Disposal";
                        list.Count = item.RateDisposal;

                        result.Add(list);
                    }
                    list = null;
                    list = new GeneratedList();
                }
                if (item.DeliveryName != null)
                {
                    // Перевіряєм чи вже є такий запис, якщо є то додаєм тільки кількість. Не робим дублів. 
                    var totalResult = result.FirstOrDefault(res => res.GeneratedId == select.Id &&
                                                                   res.Groupe == "FLOORING DELIVERY" &&
                                                                   res.Name == item.DeliveryName);
                    if (totalResult != null)
                    {
                        int index = result.IndexOf(totalResult);
                        if (index != -1)
                        {
                            result[index].Count += item.DeliveryQty;
                        }
                    }
                    else
                    {
                        list.GeneratedId = select.Id;
                        list.Groupe = "FLOORING DELIVERY";
                        list.Name = item.DeliveryName;
                        list.Count = item.DeliveryQty;

                        result.Add(list);
                    }
                    list = null;
                    list = new GeneratedList();
                }
            }
            
            int countFlood = 0;
            foreach (var item in flood)
            {
                if (item.Depth != null)
                {
                    // Перевіряєм чи вже є такий запис, якщо є то додаєм тільки кількість. Не робим дублів. 
                    var totalResult = result.FirstOrDefault(res => res.GeneratedId == select.Id &&
                                                                   res.Groupe == "FLOORING" &&
                                                                   res.Name == item.Concatenar);
                    if (totalResult != null)
                    {
                        int index = result.IndexOf(totalResult);
                        if (index != -1)
                        {
                            result[index].Count += item.QtyNovoplan;
                        }
                    }
                    else
                    {
                        list.GeneratedId = select.Id;
                        list.Groupe = "FLOORING";
                        list.Name = item.Concatenar;
                        list.Count = item.QtyNovoplan;

                        result.Add(list);
                    }
                    list = null;
                    list = new GeneratedList();

                    countFlood++;
                }
            }
            if (countFlood > 0)
            {
                
                list.GeneratedId = select.Id;
                list.Groupe = "OPTIONAL SERVICES";
                list.Name = "Leveling Services Fee";
                list.Count = countFlood;

                result.Add(list);
               
                list = null;
                list = new GeneratedList();
            }

            db.GeneratedLists.AddRange(SortedListGenerated(result));
            db.SaveChanges();

            return db.GeneratedLists.Where(g => g.GeneratedId == select.Id).ToList();
        }
        private List<GeneratedList> SortedListGenerated(List<GeneratedList> list)
        {
            var flooring = list.Where(l => l.Groupe == "FLOORING").ToList();
            var accessories = list.Where(l => l.Groupe == "ACCESSORIES").ToList();
            var installation = list.Where(l => l.Groupe == "INSTALLATION").ToList();
            var demolition = list.Where(l => l.Groupe == "DEMOLITION").ToList();
            var services = list.Where(l => l.Groupe == "OPTIONAL SERVICES").ToList();
            var delivery = list.Where(l => l.Groupe == "FLOORING DELIVERY").ToList();

            List<GeneratedList> sorted = new List<GeneratedList>();

            foreach (var item in flooring)
            {
                sorted.Add(item);                 
            }
            foreach (var item in accessories)
            {
                sorted.Add(item);
            }
            foreach (var item in installation)
            {
                sorted.Add(item);
            }
            foreach (var item in demolition)
            {
                sorted.Add(item);
            }
            foreach (var item in services)
            {
                sorted.Add(item);
            }
            foreach (var item in delivery)
            {
                sorted.Add(item);
            }
            return sorted;
        }
        private Quotation AddMaterialQuotaToGenerated(List<GeneratedList> generated, Quotation select, string optionQuota)
        {
            Quotation quota = select;
            var oldMaterial = db.MaterialQuotations.Where(m => m.QuotationId == quota.Id);
            db.MaterialQuotations.RemoveRange(oldMaterial);
            db.SaveChanges();

            List<MaterialQuotation> materials = new List<MaterialQuotation>();
            List<GeneratedList> result = new List<GeneratedList>();

            if (optionQuota == "Material")
            {
                result = generated.Where(r=>r.Groupe == "FLOORING" || r.Groupe == "ACCESSORIES").ToList();
            }
            if (optionQuota == "Labour")
            {
                //result = generated.Where(r => r.Groupe != "FLOORING" || r.Groupe != "ACCESSORIES").ToList();
                result = generated.Where(r => r.Groupe == "INSTALLATION" ||
                                              r.Groupe == "DEMOLITION" ||
                                              r.Groupe == "OPTIONAL SERVICES" ||
                                              r.Groupe == "FLOORING DELIVERY").ToList();
            }
            if(optionQuota == "All")
            {
                result = generated;
            }

            foreach (var item in result)
            {
                materials.Add(new MaterialQuotation() 
                {
                    GradeLevel = item.GradeLevel,
                    Aditional = item.Aditional,
                    Partition = item.Partition,
                    Groupe = item.Groupe,                    
                    QuotationId = quota.Id,
                    Quantity = item.Count,
                    Description = item.Name,
                    GeneratedId = item.GeneratedId                    
                });
            }
            db.MaterialQuotations.AddRange(materials);
            db.SaveChanges();
            return quota;
        }
    }
}
