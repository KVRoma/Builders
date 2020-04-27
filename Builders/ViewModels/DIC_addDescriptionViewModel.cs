using Builders.Commands;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.ViewModels
{
    public class DIC_addDescriptionViewModel : ViewModel
    {
        public string WindowName { get; } = "Tools Description";
        private string nameDescription;
        private decimal price;
        private decimal cost;        
        private bool pressButton;
        

        public string NameDescription
        {
            get { return nameDescription; }
            set
            {
                nameDescription = value;
                OnPropertyChanged(nameof(NameDescription));
            }
        }
        public decimal Price
        {
            get { return price; }
            set
            {
                price = value;
                OnPropertyChanged(nameof(Price));
            }
        }
        public decimal Cost
        {
            get { return cost; }
            set 
            {
                cost = value;
                OnPropertyChanged(nameof(Cost));
            }
        }        
        public bool PressButton
        {
            get { return pressButton; }
            set
            {
                pressButton = value;
                OnPropertyChanged(nameof(PressButton));
            }
        }

        private Command _saveCommand;

        public Command SaveCommand => _saveCommand ?? (_saveCommand = new Command(obj =>
        {           

            if (NameDescription != "" )
            {
                if (obj is System.Windows.Window)
                {
                    PressButton = true;
                    (obj as System.Windows.Window).Close();
                }
            }
        }));

        public DIC_addDescriptionViewModel()
        {
            PressButton = false;            
        }
    }
}
