using Builders.Commands;
using Builders.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Data.Entity;

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
                    FloorTypeAccessories = db.DIC_G_TypeAccessoriesFloors.Where(t => t.DIC_G_AccessoriesFloorId == FloorNameAccessorieSelect.Id).ToList();
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
                if (GeneratedAccessorieSelect != null)
                {
                    TotalMaterialAccessorie = GeneratedAccessorieSelect.NewFloor;
                }
                else 
                {
                    TotalMaterialAccessorie = null;
                }
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

        #region Material Detail Command
        private Command _addDetailCommand;
        //private Command _insDetailCommand;
        private Command _delDetailCommand;
        private Command _clearDetailCommand;

        public Command AddDetailCommand => _addDetailCommand ?? (_addDetailCommand = new Command(obj =>
        {
            GeneratedMaterial material = new GeneratedMaterial()
            {
                Aditional = AdditionalDetailSelect?.Name,
                ExistingFloor = ExistingDetailSelect?.Name,
                GeneratedId = generatedSelect.Id,
                GradeLevel = LevelDetailSelect?.Name,
                LenghtFloor = LenghtDetail,
                NeedDemolition = DemolitionDetailSelect,
                NewFloor = NewFloorDetailSelect?.Name,
                NoteTransitionOther = TransitionOtherNameDetail,
                Partition = PartitionDetailSelect?.Name,
                QtyTrim = TrimQtyDetail,
                TotalFloor = decimal.Round(WidthDetail * LenghtDetail, 2),
                TotalFlooringMaterial = decimal.Round((WidthDetail * LenghtDetail) * 1.1m, 2),
                TransitionOther = TransitionOtherDetail,
                TransitionR = TransitionRDetail,
                TransitionT = TransitionTDetail,
                TypeTrim = TrimDetailSelect?.Name,
                WidthFloor = WidthDetail
            };
            db.GeneratedMaterials.Add(material);
            db.SaveChanges();
            CalculateAccessories(material, true);
            LoadDetail();            
        }));
        //public Command InsDetailCommand => _insDetailCommand ?? (_insDetailCommand = new Command(obj =>
        //{
        //    if (GeneratedMaterialSelect != null)
        //    {
        //        GeneratedMaterialSelect.Aditional = AdditionalDetailSelect?.Name;
        //        GeneratedMaterialSelect.ExistingFloor = ExistingDetailSelect?.Name;
        //        GeneratedMaterialSelect.GeneratedId = generatedSelect.Id;
        //        GeneratedMaterialSelect.GradeLevel = LevelDetailSelect?.Name;
        //        GeneratedMaterialSelect.LenghtFloor = LenghtDetail;
        //        GeneratedMaterialSelect.NeedDemolition = DemolitionDetailSelect;
        //        GeneratedMaterialSelect.NewFloor = NewFloorDetailSelect?.Name;
        //        GeneratedMaterialSelect.NoteTransitionOther = TransitionOtherNameDetail;
        //        GeneratedMaterialSelect.Partition = PartitionDetailSelect?.Name;
        //        GeneratedMaterialSelect.QtyTrim = TrimQtyDetail;
        //        GeneratedMaterialSelect.TotalFloor = decimal.Round(WidthDetail * LenghtDetail, 2);
        //        GeneratedMaterialSelect.TotalFlooringMaterial = decimal.Round((WidthDetail * LenghtDetail) * 1.1m, 2);
        //        GeneratedMaterialSelect.TransitionOther = TransitionOtherDetail;
        //        GeneratedMaterialSelect.TransitionR = TransitionRDetail;
        //        GeneratedMaterialSelect.TransitionT = TransitionTDetail;
        //        GeneratedMaterialSelect.TypeTrim = TrimDetailSelect?.Name;
        //        GeneratedMaterialSelect.WidthFloor = WidthDetail;
        //    }
        //    db.Entry(GeneratedMaterialSelect).State = EntityState.Modified;
        //    db.SaveChanges();
        //    CalculateAccessories(GeneratedMaterialSelect, true);
        //    LoadDetail();            
        //}));
        public Command DelDetailCommand => _delDetailCommand ?? (_delDetailCommand = new Command(obj =>
        {
            if (GeneratedMaterialSelect != null)
            {
                CalculateAccessories(GeneratedMaterialSelect, false);
                db.GeneratedMaterials.Remove(GeneratedMaterialSelect);
                db.SaveChanges();
                LoadDetail();
            }
        }));
        public Command ClearDetailCommand => _clearDetailCommand ?? (_clearDetailCommand = new Command(obj =>
        {
            LevelDetailSelect = null;
            PartitionDetailSelect = null;
            AdditionalDetailSelect = null;
            TransitionRDetail = 0m;
            TransitionTDetail = 0m;
            TransitionOtherDetail = 0m;
            TransitionOtherNameDetail = null;
            TrimDetailSelect = null;
            TrimQtyDetail = 0m;
            ExistingDetailSelect = null;
            NewFloorDetailSelect = null;
            DemolitionDetailSelect = DemolitionDetails[1];
            LenghtDetail = 0m;
            WidthDetail = 0m;
        }));
        #endregion
        #region Accessories Command
        private Command _insAccessoriesCommand;

        public Command InsAccessoriesCommand => _insAccessoriesCommand ?? (_insAccessoriesCommand = new Command(obj=> 
        {
            GeneratedAccessorieSelect.AccessoriesFloor = FloorNameAccessorieSelect.Name;
            GeneratedAccessorieSelect.TypeAccessoriesFloor = FloorTypeAccessorieSelect.Name;
            db.Entry(GeneratedAccessorieSelect).State = EntityState.Modified;
            db.SaveChanges();
            LoadAccessoriess();
        }));
        #endregion
        #region Stairs Command
        private Command _addStairsCommand;
        private Command _insStairsCommand;
        private Command _delStairsCommand;
        private Command _clearStairsCommand;

        public Command AddStairsCommand => _addStairsCommand ?? (_addStairsCommand = new Command(obj =>
        {
            decimal qty = 0m;
            if (TypeStairSelect?.Name == "Overlap Nosing" ||
                TypeStairSelect?.Name == "Regular Nosing" ||
                TypeStairSelect?.Name == "Overlap Nosing - over size" ||
                TypeStairSelect?.Name == "Regular Nosing - over size")
            {
                qty = decimal.Round(QtyStairs * LenghtStairs, 2);
            }
            else 
            {
                qty = QtyStairs;
            }

            GeneratedStairs stairs = new GeneratedStairs()
            {
                Demolition = DemolitionStairSelect,
                GeneratedId = generatedSelect.Id,
                GradeLevel = LevelStairSelect?.Name,
                LenghtStairs = LenghtStairs,
                QtyLeveling = QtyLevelingStairs,
                QtyStairs = QtyStairs,
                QtyStairsLenght = qty,
                TypeLeveling = TypeLevelingStairSelect?.Name,
                TypeStairs = TypeStairSelect?.Name                
            };
            db.GeneratedStairs.Add(stairs);
            db.SaveChanges();
            LoadStairs();
        }));
        public Command InsStairsCommand => _insStairsCommand ?? (_insStairsCommand = new Command(obj=> 
        {
            if (GeneratedStairSelect != null)
            {
                decimal qty = 0m;
                if (TypeStairSelect.Name == "Overlap Nosing" ||
                    TypeStairSelect.Name == "Regular Nosing" ||
                    TypeStairSelect.Name == "Overlap Nosing - over size" ||
                    TypeStairSelect.Name == "Regular Nosing - over size")
                {
                    qty = decimal.Round(QtyStairs * LenghtStairs, 2);
                }
                else
                {
                    qty = QtyStairs;
                }

                GeneratedStairSelect.Demolition = DemolitionStairSelect;
                GeneratedStairSelect.GeneratedId = generatedSelect.Id;
                GeneratedStairSelect.GradeLevel = LevelStairSelect?.Name;
                GeneratedStairSelect.LenghtStairs = LenghtStairs;
                GeneratedStairSelect.QtyLeveling = QtyLevelingStairs;
                GeneratedStairSelect.QtyStairs = QtyStairs;
                GeneratedStairSelect.QtyStairsLenght = qty;
                GeneratedStairSelect.TypeLeveling = TypeLevelingStairSelect?.Name;
                GeneratedStairSelect.TypeStairs = TypeStairSelect?.Name;
                db.Entry(GeneratedStairSelect).State = EntityState.Modified;
                db.SaveChanges();
                LoadStairs();
            }
        }));
        public Command DelStairsCommand => _delStairsCommand ?? (_delStairsCommand = new Command(obj=> 
        {
            if (GeneratedStairSelect != null)
            {
                db.GeneratedStairs.Remove(GeneratedStairSelect);
                db.SaveChanges();
                LoadStairs();
            }
        }));
        public Command ClearStairsCommand => _clearStairsCommand ?? (_clearStairsCommand = new Command(obj=> 
        {
            LevelStairSelect = null;
            TypeStairSelect = null;
            QtyStairs = 0m;
            LenghtStairs = 0m;
            DemolitionStairSelect = DemolitionStairs[1];
            TypeLevelingStairSelect = null;
            QtyLevelingStairs = 0m;
        }));
        #endregion
        #region Moldings Command
        private Command _addMoldingsCommand;
        private Command _insMoldingsCommand;
        private Command _delMoldingsCommand;
        private Command _clearMoldingsCommand;

        public Command AddMoldingsCommand => _addMoldingsCommand ?? (_addMoldingsCommand = new Command(obj=> 
        {
            decimal qty = 0m;
            if (AccessoriesMoldingSelect?.Name == "Baseboards")
            {
                qty = (decimal.Ceiling(QtyMolding / 16m) * 16m) + 16m;
            }
            else if (AccessoriesMoldingSelect?.Name == "Door Casing - Over Size")
            {
                qty = decimal.Ceiling(QtyMolding * 1.5m) * 16m;
            }
            else if (AccessoriesMoldingSelect?.Name == "Door Casing")
            {
                qty = decimal.Ceiling(QtyMolding * 1.25m) * 16m;
            }
            GeneratedMolding molding = new GeneratedMolding() 
            {
                BaseboardMaterial = qty,
                GeneratedId = generatedSelect.Id,
                HeightMolding = HeightMolding,
                MoldingName = AccessoriesMoldingSelect?.Name,
                Painting = PaintingMoldingSelect?.Name,
                QtyMolding = QtyMolding,
                TypeMolding = TypeMoldingSelect?.Name                
            };
            db.GeneratedMoldings.Add(molding);
            db.SaveChanges();
            LoadMolding();
        }));
        public Command InsMoldingsCommand => _insMoldingsCommand ?? (_insMoldingsCommand = new Command(obj=> 
        {
            if (GeneratedMoldingSelect != null)
            {
                decimal qty = 0m;
                if (AccessoriesMoldingSelect?.Name == "Baseboards")
                {
                    qty = (decimal.Ceiling(QtyMolding / 16m) * 16m) + 16m;
                }
                else if (AccessoriesMoldingSelect?.Name == "Door Casing - Over Size")
                {
                    qty = decimal.Ceiling(QtyMolding * 1.5m) * 16m;
                }
                else if (AccessoriesMoldingSelect?.Name == "Door Casing")
                {
                    qty = decimal.Ceiling(QtyMolding * 1.25m) * 16m;
                }

                GeneratedMoldingSelect.BaseboardMaterial = qty;
                GeneratedMoldingSelect.GeneratedId = generatedSelect.Id;
                GeneratedMoldingSelect.HeightMolding = HeightMolding;
                GeneratedMoldingSelect.MoldingName = AccessoriesMoldingSelect?.Name;
                GeneratedMoldingSelect.Painting = PaintingMoldingSelect?.Name;
                GeneratedMoldingSelect.QtyMolding = QtyMolding;
                GeneratedMoldingSelect.TypeMolding = TypeMoldingSelect?.Name;

                db.Entry(GeneratedMoldingSelect).State = EntityState.Modified;
                db.SaveChanges();
                LoadMolding();
            }
        }));
        public Command DelMoldingsCommand => _delMoldingsCommand ?? (_delMoldingsCommand = new Command(obj=> 
        {
            if (GeneratedMoldingSelect != null)
            {
                db.GeneratedMoldings.Remove(GeneratedMoldingSelect);
                db.SaveChanges();
                LoadMolding();
            }
        }));
        public Command ClearMoldingsCommand => _clearMoldingsCommand ?? (_clearMoldingsCommand = new Command(obj=>
        {
            AccessoriesMoldingSelect = null;
            TypeMoldingSelect = null;
            QtyMolding = 0m;
            HeightMolding = 0m;
            PaintingMoldingSelect = null;
        }));
        #endregion
        #region Suplementary Command
        private Command _addSuplementaryCommand;
        private Command _insSuplementaryCommand;
        private Command _delSuplementaryCommand;
        private Command _clearSuplementaryCommand;

        public Command AddSuplementaryCommand => _addSuplementaryCommand ?? (_addSuplementaryCommand = new Command(obj=> 
        {
            GeneratedSuplementary suplementary = new GeneratedSuplementary()
            {
                DeliveryName = DeliverySuplementarySelect,
                DeliveryQty = QtyDeliverySuplementary,
                FurnitureHandelingRoom = RoomSuplementary,
                GeneralNotes = NotesSuplementary,
                GeneratedId = generatedSelect.Id,
                RateDisposal = DisposalSuplementary
            };
            db.GeneratedSuplementaries.Add(suplementary);
            db.SaveChanges();
            LoadSuplementary();
        }));
        public Command InsSuplementaryCommand => _insSuplementaryCommand ?? (_insSuplementaryCommand = new Command(obj=> 
        {
            if (GeneratedSuplementarySelect != null)
            {
                GeneratedSuplementarySelect.DeliveryName = DeliverySuplementarySelect;
                GeneratedSuplementarySelect.DeliveryQty = QtyDeliverySuplementary;
                GeneratedSuplementarySelect.FurnitureHandelingRoom = RoomSuplementary;
                GeneratedSuplementarySelect.GeneralNotes = NotesSuplementary;
                GeneratedSuplementarySelect.GeneratedId = generatedSelect.Id;
                GeneratedSuplementarySelect.RateDisposal = DisposalSuplementary;

                db.Entry(GeneratedSuplementarySelect).State = EntityState.Modified;
                db.SaveChanges();
                LoadSuplementary();
            }
        }));
        public Command DelSuplementaryCommand => _delSuplementaryCommand ?? (_delSuplementaryCommand = new Command(obj=> 
        {
            if (GeneratedSuplementarySelect != null)
            {
                db.GeneratedSuplementaries.Remove(GeneratedSuplementarySelect);
                db.SaveChanges();
                LoadSuplementary();
            }
        }));
        public Command ClearSuplementaryCommand => _clearSuplementaryCommand ?? (_clearSuplementaryCommand = new Command(obj=> 
        {
            RoomSuplementary = 0m;
            DisposalSuplementary = 0m;
            DeliverySuplementarySelect = null;
            QtyDeliverySuplementary = 0m;
            NotesSuplementary = null;
        }));
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
        /// <summary>
        /// Завантажує якщо є, або створює Generator для заданої Quota
        /// </summary>
        /// <param name="id"></param>
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
        /// <summary>
        /// Завантажує дані в ComboBox
        /// </summary>
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
        /// <summary>
        /// Завантажує GeneratedMaterials
        /// </summary>
        private void LoadDetail()
        {
            GeneratedMaterials = null;
            GeneratedMaterials = db.GeneratedMaterials.Where(g => g.GeneratedId == generatedSelect.Id).ToList();
        }
        /// <summary>
        /// Завантажує GeneratedAccessories
        /// </summary>
        private void LoadAccessoriess()
        {
            GeneratedAccessories = null;
            GeneratedAccessories = db.GeneratedAccessories.Where(g => g.GeneratedId == generatedSelect.Id).ToList();
        }
        /// <summary>
        /// Завантажує GeneratedStairs
        /// </summary>
        private void LoadStairs()
        {
            GeneratedStairs = null;
            GeneratedStairs = db.GeneratedStairs.Where(g => g.GeneratedId == generatedSelect.Id).ToList();
        }
        /// <summary>
        /// Завантажує GeneratedMoldings
        /// </summary>
        private void LoadMolding()
        {
            GeneratedMoldings = null;
            GeneratedMoldings = db.GeneratedMoldings.Where(g => g.GeneratedId == generatedSelect.Id).ToList();
        }
        /// <summary>
        /// Завантажує GeneratedSuplementarys
        /// </summary>
        private void LoadSuplementary()
        {
            GeneratedSuplementarys = null;
            GeneratedSuplementarys = db.GeneratedSuplementaries.Where(g => g.GeneratedId == generatedSelect.Id).ToList();
        }
        /// <summary>
        /// Завантажує GeneratedFloods
        /// </summary>
        private void LoadFlood()
        {
            GeneratedFloods = null;
            GeneratedFloods = db.GeneratedFloods.Where(g => g.GeneratedId == generatedSelect.Id).ToList();
        }
        /// <summary>
        /// Округляє до верхньої ближньої сотні (264 -> 300)
        /// </summary>
        /// <param name="number"></param>
        /// <returns></returns>
        private decimal RoundToHundreds(decimal number)
        {
            decimal result = 0m;
            if (number < 100)
            {
                result = decimal.Ceiling(number / 10) * 10m;
            }
            else if (number >= 100)
            {
                result = decimal.Ceiling(number / 100) * 100m;
            }
            
            return result;
        }
        /// <summary>
        /// Перераховує, або створює назву та кількість NewFloor в Accessories 
        /// </summary>
        /// <param name="generatedMaterial">Вибраний екземпляр Detail</param>
        /// <param name="key">true - додати елемент, false - відняти</param>
        private void CalculateAccessories(GeneratedMaterial generatedMaterial, bool key)
        {
            GeneratedAccessories accessories = GeneratedAccessories.FirstOrDefault(a => a.NewFloor == generatedMaterial.NewFloor);
            if (key)
            {
                if (accessories != null)
                {
                    accessories.TotalFloor += generatedMaterial?.TotalFloor ?? 0m;
                    accessories.TotalAccessoriesFloor = RoundToHundreds(accessories?.TotalFloor ?? 0m);
                    db.Entry(accessories).State = EntityState.Modified;
                    db.SaveChanges();
                    LoadAccessoriess();
                }
                else
                {
                    if (generatedMaterial.NewFloor != null)
                    {
                        GeneratedAccessories accessorie = new GeneratedAccessories()
                        {
                            NewFloor = generatedMaterial.NewFloor,
                            TotalFloor = generatedMaterial?.TotalFloor ?? 0m,
                            TotalAccessoriesFloor = RoundToHundreds(generatedMaterial?.TotalFloor ?? 0m),
                            GeneratedId = generatedMaterial.GeneratedId
                        };
                        db.GeneratedAccessories.Add(accessorie);
                        db.SaveChanges();
                        LoadAccessoriess();
                    }
                }
            }
            else 
            {
                if (accessories != null)
                {
                    accessories.TotalFloor -= generatedMaterial?.TotalFloor ?? 0m;
                    if (accessories.TotalFloor != 0)
                    {
                        accessories.TotalAccessoriesFloor = RoundToHundreds(accessories?.TotalFloor ?? 0m);
                        db.Entry(accessories).State = EntityState.Modified;
                    }
                    else
                    {
                        db.GeneratedAccessories.Remove(accessories);
                    }
                    db.SaveChanges();
                    LoadAccessoriess();
                }
            }
            //else
            //{
            //    var access = GeneratedAccessories.Where(a => a.GeneratedId == generatedMaterial.GeneratedId);
            //    foreach(var item in access)
            //    {
            //        var res = db.GeneratedMaterials.Where(a => a.GeneratedId == generatedMaterial.GeneratedId && a.NewFloor == item.NewFloor);
            //        if (res != null)
            //        {
            //            item.TotalFloor = res.Select(r => r.TotalFloor)?.Sum() ?? 0m;
            //            item.TotalAccessoriesFloor = RoundToHundreds(res.Select(r=>r.TotalFloor)?.Sum() ?? 0m);
            //            db.Entry(item).State = EntityState.Modified;
            //        }
            //        else
            //        {
            //            db.GeneratedAccessories.Remove(item);
            //        }
            //    }
            //    db.SaveChanges();
            
            LoadAccessoriess();
        }
    }
}
