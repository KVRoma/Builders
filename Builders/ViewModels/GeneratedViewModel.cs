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

        #region Material Detail Properti
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


        public GeneratedViewModel(ref BuilderContext context, int id)
        {
            db = context;            

            LoadGenerator(id);
            LoadDetail();
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
        private void LoadDetail()
        {
            GeneratedMaterials = db.GeneratedMaterials.Where(g=>g.GeneratedId == generatedSelect.Id).ToList();
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
        }
    }
}
