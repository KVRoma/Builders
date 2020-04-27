using Builders.Commands;
using Builders.Models;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using Color = System.Drawing.Color;

namespace Builders.ViewModels
{
    public class DIC_ContractorViewModel : ViewModel
    {
        private BuilderContext db;
        public string WindowName { get; } = "Contractors";

        private DIC_Contractor contractorSelect;
        private IEnumerable<DIC_Contractor> contractors;
        private string colourSelect;
        private IList<string> colour;
       

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
        public string ColourSelect
        {
            get { return colourSelect; }
            set
            {
                colourSelect = value;
                OnPropertyChanged(nameof(ColourSelect));
            }
        }
        public IList<string> Colour
        {
            get { return colour; }
            set
            {
                colour = value;
                OnPropertyChanged(nameof(Colour));
            }
        }

        private Command addCommand;        
        private Command delCommand;

        public Command AddCommand => addCommand ?? (addCommand = new Command(obj=> 
        {
            db.SaveChanges();
            
        }));
       
        public Command DelCommand => delCommand ?? (delCommand = new Command(obj=> 
        {
            if (ContractorSelect != null && ContractorSelect.Id != 1)
            {
                db.DIC_Contractors.Remove(ContractorSelect);
                db.SaveChanges();
                
            }
        }));

        public DIC_ContractorViewModel(ref BuilderContext basa)
        {
            db = basa;
            db.DIC_Contractors.Load();            
            Contractors = db.DIC_Contractors.Local.ToBindingList();
            ColourLoad();
        }
       
        
        private void ColourLoad()
        {
            Colour = new List<string>();            
            string[] colorNames = Enum.GetNames(typeof(KnownColor));            
            foreach (string colorName in colorNames)
            {                
                KnownColor knownColor = (KnownColor)Enum.Parse(typeof(KnownColor), colorName);               
                if (knownColor > KnownColor.Transparent)
                {                  
                   Colour.Add(colorName);
                }
            }
        }
    }
}
