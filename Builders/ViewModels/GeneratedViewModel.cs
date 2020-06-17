using Builders.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.ViewModels
{
    public class GeneratedViewModel : ViewModel
    {
        public string WindowName { get; set; } = "Generator";
        private BuilderContext db;        
        private Generated generatedSelect;

        #region Material Detail Property
        private DIC_G_GradeLevel levelDetailSelect;
        private List<DIC_G_GradeLevel> levelDetails;
        private DIC_G_Partition partitionDetailSelect;
        private List<DIC_G_Partition> partitionDetails;
        private DIC_G_Additional additionalDetailSelect;
        private List<DIC_G_Additional> additionalDetails;
        private decimal transitionRDetail;
        private decimal transitionTDetail;
        private decimal transitionOtherDetail;
        private string transitionOtherNameDetail;
        private DIC_G_TypeTrim trimDetailSelect;
        private List<DIC_G_TypeTrim> trimDetails;
        private decimal trimQtyDetail;
        private DIC_G_Floor existingDetailSelect;
        private List<DIC_G_Floor> existingDetails;
        private DIC_G_Floor newFloorDetailSelect;
        private List<DIC_G_Floor> newFloorDetails;
        private string demolitionDetailSelect;
        private List<string> demolitionDetails;
        private decimal lenghtDetail;
        private decimal widthDetail;
        private GeneratedMaterial generatedMaterialSelect;
        private List<GeneratedMaterial> generatedMaterials;

        public DIC_G_GradeLevel LevelDetailSelect
        {
            get => levelDetailSelect; 
            set
            {
                levelDetailSelect = value;
                OnPropertyChanged(nameof(LevelDetailSelect));
            }
        }
        public List<DIC_G_GradeLevel> LevelDetails
        {
            get => levelDetails; 
            set
            {
                levelDetails = value;
                OnPropertyChanged(nameof(LevelDetails));
            }
        }
        public DIC_G_Partition PartitionDetailSelect
        {
            get => partitionDetailSelect; 
            set
            {
                partitionDetailSelect = value;
                OnPropertyChanged(nameof(PartitionDetailSelect));
            }
        }
        public List<DIC_G_Partition> PartitionDetails
        {
            get => partitionDetails; 
            set
            {
                partitionDetails = value;
                OnPropertyChanged(nameof(PartitionDetails));
            }
        }
        public DIC_G_Additional AdditionalDetailSelect
        {
            get => additionalDetailSelect; 
            set
            {
                additionalDetailSelect = value;
                OnPropertyChanged(nameof(AdditionalDetailSelect));
            }
        }
        public List<DIC_G_Additional> AdditionalDetails
        {
            get => additionalDetails; 
            set
            {
                additionalDetails = value;
                OnPropertyChanged(nameof(AdditionalDetails));
            }
        }
        public decimal TransitionRDetail
        {
            get => transitionRDetail; 
            set
            {
                transitionRDetail = value;
                OnPropertyChanged(nameof(TransitionRDetail));
            }
        }
        public decimal TransitionTDetail
        {
            get => transitionTDetail; 
            set
            {
                transitionTDetail = value;
                OnPropertyChanged(nameof(TransitionTDetail));
            }
        }
        public decimal TransitionOtherDetail
        {
            get => transitionOtherDetail; 
            set
            {
                transitionOtherDetail = value;
                OnPropertyChanged(nameof(TransitionOtherDetail));
            }
        }
        public string TransitionOtherNameDetail
        {
            get => transitionOtherNameDetail; 
            set
            {
                transitionOtherNameDetail = value;
                OnPropertyChanged(nameof(TransitionOtherNameDetail));
            }
        }
        public DIC_G_TypeTrim TrimDetailSelect
        {
            get => trimDetailSelect; 
            set
            {
                trimDetailSelect = value;
                OnPropertyChanged(nameof(TrimDetailSelect));
            }
        }
        public List<DIC_G_TypeTrim> TrimDetails
        {
            get => trimDetails; 
            set
            {
                trimDetails = value;
                OnPropertyChanged(nameof(TrimDetails));
            }
        }
        public decimal TrimQtyDetail
        {
            get => trimQtyDetail; 
            set
            {
                trimQtyDetail = value;
                OnPropertyChanged(nameof(TrimQtyDetail));
            }
        }
        public DIC_G_Floor ExistingDetailSelect
        {
            get => existingDetailSelect; 
            set
            {
                existingDetailSelect = value;
                OnPropertyChanged(nameof(ExistingDetailSelect));
            }
        }
        public List<DIC_G_Floor> ExistingDetails
        {
            get => existingDetails; 
            set
            {
                existingDetails = value;
                OnPropertyChanged(nameof(ExistingDetails));
            }
        }
        public DIC_G_Floor NewFloorDetailSelect
        {
            get => newFloorDetailSelect; 
            set
            {
                newFloorDetailSelect = value;
                OnPropertyChanged(nameof(NewFloorDetailSelect));
            }
        }
        public List<DIC_G_Floor> NewFloorDetails
        {
            get => newFloorDetails;
            set
            {
                newFloorDetails = value;
                OnPropertyChanged(nameof(NewFloorDetails));
            }
        }
        public string DemolitionDetailSelect
        {
            get => demolitionDetailSelect; 
            set
            {
                demolitionDetailSelect = value;
                OnPropertyChanged(nameof(DemolitionDetailSelect));
            }
        }
        public List<string> DemolitionDetails
        {
            get => demolitionDetails; 
            set
            {
                demolitionDetails = value;
                OnPropertyChanged(nameof(DemolitionDetails));
            }
        }
        public decimal LenghtDetail
        {
            get => lenghtDetail; 
            set
            {
                lenghtDetail = value;
                OnPropertyChanged(nameof(LenghtDetail));
            }
        }
        public decimal WidthDetail
        {
            get => widthDetail; 
            set
            {
                widthDetail = value;
                OnPropertyChanged(nameof(WidthDetail));
            }
        }
        public GeneratedMaterial GeneratedMaterialSelect
        {
            get => generatedMaterialSelect; 
            set
            {
                generatedMaterialSelect = value;
                OnPropertyChanged(nameof(GeneratedMaterialSelect));
            }
        }
        public List<GeneratedMaterial> GeneratedMaterials
        {
            get => generatedMaterials; 
            set
            {
                generatedMaterials = value;
                OnPropertyChanged(nameof(GeneratedMaterials));
            }
        }
        #endregion
        #region Material Accessories Property
        private DIC_G_AccessoriesFloor floorNameAccessorieSelect;
        private List<DIC_G_AccessoriesFloor> floorNameAccessories;
        private DIC_G_TypeAccessoriesFloor floorTypeAccessorieSelect;
        private List<DIC_G_TypeAccessoriesFloor> floorTypeAccessories;
        private string totalMaterialAccessorie;
        private GeneratedAccessories generatedAccessorieSelect;
        private List<GeneratedAccessories> generatedAccessories;

        public DIC_G_AccessoriesFloor FloorNameAccessorieSelect
        {
            get => floorNameAccessorieSelect; 
            set
            {
                floorNameAccessorieSelect = value;
                OnPropertyChanged(nameof(FloorNameAccessorieSelect));
                if (FloorNameAccessorieSelect != null)
                {
                    FloorTypeAccessories = db.DIC_G_TypeAccessoriesFloors.Where(t=>t.DIC_G_AccessoriesFloorId == FloorNameAccessorieSelect.Id).ToList();
                }
            }
        }
        public List<DIC_G_AccessoriesFloor> FloorNameAccessories
        {
            get => floorNameAccessories; 
            set
            {
                floorNameAccessories = value;
                OnPropertyChanged(nameof(FloorNameAccessories));
            }
        }
        public DIC_G_TypeAccessoriesFloor FloorTypeAccessorieSelect
        {
            get => floorTypeAccessorieSelect; 
            set
            {
                floorTypeAccessorieSelect = value;
                OnPropertyChanged(nameof(FloorTypeAccessorieSelect));
            }
        }
        public List<DIC_G_TypeAccessoriesFloor> FloorTypeAccessories
        {
            get => floorTypeAccessories; 
            set
            {
                floorTypeAccessories = value;
                OnPropertyChanged(nameof(FloorTypeAccessories));
            }
        }
        public string TotalMaterialAccessorie
        {
            get => totalMaterialAccessorie; 
            set
            {
                totalMaterialAccessorie = value;
                OnPropertyChanged(nameof(TotalMaterialAccessorie));
            }
        }
        public GeneratedAccessories GeneratedAccessorieSelect
        {
            get => generatedAccessorieSelect; 
            set
            {
                generatedAccessorieSelect = value;
                OnPropertyChanged(nameof(GeneratedAccessorieSelect));
            }
        }
        public List<GeneratedAccessories> GeneratedAccessories
        {
            get => generatedAccessories; 
            set
            {
                generatedAccessories = value;
                OnPropertyChanged(nameof(GeneratedAccessories));
            }
        }                
        #endregion        
        #region Stairs Property
        private DIC_G_GradeLevel levelStairSelect;
        private List<DIC_G_GradeLevel> levelStairs;
        private DIC_G_TypeStairs typeStairSelect;
        private List<DIC_G_TypeStairs> typeStairs;
        private decimal qtyStairs;
        private decimal lenghtStairs;
        private string demolitionStairSelect;
        private List<string> demolitionStairs;
        private DIC_G_TypeLeveling typeLevelingStairSelect;
        private List<DIC_G_TypeLeveling> typeLevelingStairs;
        private decimal qtyLevelingStairs;
        private GeneratedStairs generatedStairSelect;
        private List<GeneratedStairs> generatedStairs;

        public DIC_G_GradeLevel LevelStairSelect
        {
            get => levelStairSelect; 
            set
            {
                levelStairSelect = value;
                OnPropertyChanged(nameof(LevelStairSelect));
            }
        }
        public List<DIC_G_GradeLevel> LevelStairs
        {
            get => levelStairs; 
            set
            {
                levelStairs = value;
                OnPropertyChanged(nameof(LevelStairs));
            }
        }
        public DIC_G_TypeStairs TypeStairSelect
        {
            get => typeStairSelect; 
            set
            {
                typeStairSelect = value;
                OnPropertyChanged(nameof(TypeStairSelect));
            }
        }
        public List<DIC_G_TypeStairs> TypeStairs
        {
            get => typeStairs; 
            set
            {
                typeStairs = value;
                OnPropertyChanged(nameof(TypeStairs));
            }
        }
        public decimal QtyStairs
        {
            get => qtyStairs; 
            set
            {
                qtyStairs = value;
                OnPropertyChanged(nameof(QtyStairs));
            }
        }
        public decimal LenghtStairs
        {
            get => lenghtStairs; 
            set
            {
                lenghtStairs = value;
                OnPropertyChanged(nameof(LenghtStairs));
            }
        }
        public string DemolitionStairSelect
        {
            get => demolitionStairSelect; 
            set
            {
                demolitionStairSelect = value;
                OnPropertyChanged(nameof(DemolitionStairSelect));
            }
        }
        public List<string> DemolitionStairs
        {
            get => demolitionStairs; 
            set
            {
                demolitionStairs = value;
                OnPropertyChanged(nameof(DemolitionStairs));
            }
        }
        public DIC_G_TypeLeveling TypeLevelingStairSelect
        {
            get => typeLevelingStairSelect; 
            set
            {
                typeLevelingStairSelect = value;
                OnPropertyChanged(nameof(TypeLevelingStairSelect));
            }
        }
        public List<DIC_G_TypeLeveling> TypeLevelingStairs
        {
            get => typeLevelingStairs; 
            set
            {
                typeLevelingStairs = value;
                OnPropertyChanged(nameof(TypeLevelingStairs));
            }
        }
        public decimal QtyLevelingStairs
        {
            get => qtyLevelingStairs; 
            set
            {
                qtyLevelingStairs = value;
                OnPropertyChanged(nameof(QtyLevelingStairs));
            }
        }
        public GeneratedStairs GeneratedStairSelect
        {
            get => generatedStairSelect; 
            set
            {
                generatedStairSelect = value;
                OnPropertyChanged(nameof(GeneratedStairSelect));
            }
        }
        public List<GeneratedStairs> GeneratedStairs
        {
            get => generatedStairs; 
            set
            {
                generatedStairs = value;
                OnPropertyChanged(nameof(GeneratedStairs));
            }
        }        
        #endregion
        #region Moldings Property
        private DIC_G_Modeling accessoriesMoldingSelect;
        private List<DIC_G_Modeling> accessoriesMoldings;
        private DIC_G_TypeModeling typeMoldingSelect;
        private List<DIC_G_TypeModeling> typeMoldings;
        private decimal qtyMolding;
        private decimal heightMolding;
        private DIC_G_Painting paintingMoldingSelect;
        private List<DIC_G_Painting> paintingMoldings;
        private GeneratedMolding generatedMoldingSelect;
        private List<GeneratedMolding> generatedMoldings;

        public DIC_G_Modeling AccessoriesMoldingSelect
        {
            get => accessoriesMoldingSelect; 
            set
            {
                accessoriesMoldingSelect = value;
                OnPropertyChanged(nameof(AccessoriesMoldingSelect));
            }
        }
        public List<DIC_G_Modeling> AccessoriesMoldings
        {
            get => accessoriesMoldings; 
            set
            {
                accessoriesMoldings = value;
                OnPropertyChanged(nameof(AccessoriesMoldings));
            }
        }
        public DIC_G_TypeModeling TypeMoldingSelect
        {
            get => typeMoldingSelect; 
            set
            {
                typeMoldingSelect = value;
                OnPropertyChanged(nameof(TypeMoldingSelect));
            }
        }
        public List<DIC_G_TypeModeling> TypeMoldings
        {
            get => typeMoldings; 
            set
            {
                typeMoldings = value;
                OnPropertyChanged(nameof(TypeMoldings));
            }
        }
        public decimal QtyMolding
        {
            get => qtyMolding; 
            set
            {
                qtyMolding = value;
                OnPropertyChanged(nameof(QtyMolding));
            }
        }
        public decimal HeightMolding
        {
            get => heightMolding; 
            set
            {
                heightMolding = value;
                OnPropertyChanged(nameof(HeightMolding));
            }
        }
        public DIC_G_Painting PaintingMoldingSelect
        {
            get => paintingMoldingSelect; 
            set
            {
                paintingMoldingSelect = value;
                OnPropertyChanged(nameof(PaintingMoldingSelect));
            }
        }
        public List<DIC_G_Painting> PaintingMoldings
        {
            get => paintingMoldings; 
            set
            {
                paintingMoldings = value;
                OnPropertyChanged(nameof(PaintingMoldings));
            }
        }
        public GeneratedMolding GeneratedMoldingSelect
        {
            get => generatedMoldingSelect; 
            set
            {
                generatedMoldingSelect = value;
                OnPropertyChanged(nameof(GeneratedMoldingSelect));
            }
        }
        public List<GeneratedMolding> GeneratedMoldings
        {
            get => generatedMoldings; 
            set
            {
                generatedMoldings = value;
                OnPropertyChanged(nameof(GeneratedMoldings));
            }
        }        
        #endregion
        #region Suplementary Property
        private decimal roomSuplementary;
        private decimal disposalSuplementary;
        private string deliverySuplementarySelect;
        private List<string> deliverySuplementarys;
        private decimal qtyDeliverySuplementary;
        private string notesSuplementary;
        private GeneratedSuplementary generatedSuplementarySelect;
        private List<GeneratedSuplementary> generatedSuplementarys;

        public decimal RoomSuplementary
        {
            get => roomSuplementary; 
            set
            {
                roomSuplementary = value;
                OnPropertyChanged(nameof(RoomSuplementary));
            }
        }
        public decimal DisposalSuplementary
        {
            get => disposalSuplementary; 
            set
            {
                disposalSuplementary = value;
                OnPropertyChanged(nameof(DisposalSuplementary));
            }
        }
        public string DeliverySuplementarySelect
        {
            get => deliverySuplementarySelect; 
            set
            {
                deliverySuplementarySelect = value;
                OnPropertyChanged(nameof(DeliverySuplementarySelect));
            }
        }
        public List<string> DeliverySuplementarys
        {
            get => deliverySuplementarys; 
            set
            {
                deliverySuplementarys = value;
                OnPropertyChanged(nameof(DeliverySuplementarys));
            }
        }
        public decimal QtyDeliverySuplementary
        {
            get => qtyDeliverySuplementary; 
            set
            {
                qtyDeliverySuplementary = value;
                OnPropertyChanged(nameof(QtyDeliverySuplementary));
            }
        }
        public string NotesSuplementary
        {
            get => notesSuplementary; 
            set
            {
                notesSuplementary = value;
                OnPropertyChanged(nameof(NotesSuplementary));
            }
        }
        public GeneratedSuplementary GeneratedSuplementarySelect
        {
            get => generatedSuplementarySelect; 
            set
            {
                generatedSuplementarySelect = value;
                OnPropertyChanged(nameof(GeneratedSuplementarySelect));
            }
        }
        public List<GeneratedSuplementary> GeneratedSuplementarys
        {
            get => generatedSuplementarys; 
            set
            {
                generatedSuplementarys = value;
                OnPropertyChanged(nameof(GeneratedSuplementarys));
            }
        }                
        #endregion
        #region Flood Property
        private string roomFlood;
        private decimal sizeFlood;
        private DIC_DepthQuotation depthFloodSelect;
        private List<DIC_DepthQuotation> depthFloods;
        private GeneratedFlood generatedFloodSelect;
        private List<GeneratedFlood> generatedFloods;

        public string RoomFlood
        {
            get => roomFlood; 
            set
            {
                roomFlood = value;
                OnPropertyChanged(nameof(RoomFlood));
            }
        }
        public decimal SizeFlood
        {
            get => sizeFlood; 
            set
            {
                sizeFlood = value;
                OnPropertyChanged(nameof(SizeFlood));
            }
        }
        public DIC_DepthQuotation DepthFloodSelect
        {
            get => depthFloodSelect; 
            set
            {
                depthFloodSelect = value;
                OnPropertyChanged(nameof(DepthFloodSelect));
            }
        }
        public List<DIC_DepthQuotation> DepthFloods
        {
            get => depthFloods; 
            set
            {
                depthFloods = value;
                OnPropertyChanged(nameof(DepthFloods));
            }
        }
        public GeneratedFlood GeneratedFloodSelect
        {
            get => generatedFloodSelect; 
            set
            {
                generatedFloodSelect = value;
                OnPropertyChanged(nameof(GeneratedFloodSelect));
            }
        }
        public List<GeneratedFlood> GeneratedFloods
        {
            get => generatedFloods; 
            set
            {
                generatedFloods = value;
                OnPropertyChanged(nameof(GeneratedFloods));
            }
        }
        #endregion

        public GeneratedViewModel(ref BuilderContext context, int id)
        {
            db = context;            

            LoadGenerator(id);
            LoadComboBox();
            LoadDetail();
            LoadAccessoriess();
            LoadStairs();
            LoadMolding();
            LoadSuplementary();
            LoadFlood();
        }

        private void LoadGenerator(int id)
        {
            generatedSelect = db.Generateds.FirstOrDefault(g => g.ClientId == id);
            if (generatedSelect == null)
            {
                generatedSelect = new Generated()
                {
                    ClientId = id
                };
                db.Generateds.Add(generatedSelect);
                db.SaveChanges();
                generatedSelect = null;
                generatedSelect = db.Generateds.FirstOrDefault(g => g.ClientId == id);
            }
        }
        private void LoadComboBox()
        {
            // Details
            LevelDetails = db.DIC_G_GradeLevels.ToList();
            PartitionDetails = db.DIC_G_Partitions.ToList();
            AdditionalDetails = db.DIC_G_Additionals.ToList();
            TrimDetails = db.DIC_G_TypeTrims.ToList();
            ExistingDetails = db.DIC_G_Floors.ToList();
            NewFloorDetails = db.DIC_G_Floors.ToList();
            DemolitionDetails = new List<string>();
            DemolitionDetails.Add("Yes");
            DemolitionDetails.Add("No");
            DemolitionDetailSelect = DemolitionDetails[1];
            // Accessories
            FloorNameAccessories = db.DIC_G_AccessoriesFloors.ToList();
            // Stairs
            LevelStairs = db.DIC_G_GradeLevels.ToList();
            TypeStairs = db.DIC_G_TypeStairs.ToList();
            DemolitionStairs = new List<string>();
            DemolitionStairs.Add("Yes");
            DemolitionStairs.Add("No");
            DemolitionStairSelect = DemolitionStairs[1];
            TypeLevelingStairs = db.DIC_G_TypeLevelings.ToList();
            // Molding
            AccessoriesMoldings = db.DIC_G_Modelings.ToList();
            TypeMoldings = db.DIC_G_TypeModelings.ToList();
            PaintingMoldings = db.DIC_G_Paintings.ToList();
            // Suplementary
            DeliverySuplementarys = new List<string>();
            var groupe = db.DIC_GroupeQuotations.FirstOrDefault(g => g.NameGroupe == "FLOORING DELIVERY");
            var items = db.DIC_ItemQuotations.Where(i => i.GroupeId == groupe.Id);
            foreach (var item in items)
            {
                var descriptions = db.DIC_DescriptionQuotations.Where(d => d.ItemId == item.Id);
                foreach (var des in descriptions)
                {
                    DeliverySuplementarys.Add(des.Name);
                }
            }
            // Depth
            DepthFloods = db.DIC_DepthQuotations.ToList();
        }
        private void LoadDetail()
        {
            GeneratedMaterials = db.GeneratedMaterials.Where(g=>g.GeneratedId == generatedSelect.Id).ToList();            
        }
        private void LoadAccessoriess()
        {
            GeneratedAccessories = db.GeneratedAccessories.Where(g => g.GeneratedId == generatedSelect.Id).ToList();
        }
        private void LoadStairs()
        {
            GeneratedStairs = db.GeneratedStairs.Where(g => g.GeneratedId == generatedSelect.Id).ToList();
        }
        private void LoadMolding()
        {
            GeneratedMoldings = db.GeneratedMoldings.Where(g => g.GeneratedId == generatedSelect.Id).ToList();
        }
        private void LoadSuplementary()
        {
            GeneratedSuplementarys = db.GeneratedSuplementaries.Where(g => g.GeneratedId == generatedSelect.Id).ToList();
        }
        private void LoadFlood()
        {
            GeneratedFloods = db.GeneratedFloods.Where(g => g.GeneratedId == generatedSelect.Id).ToList();
        }
    }
}
