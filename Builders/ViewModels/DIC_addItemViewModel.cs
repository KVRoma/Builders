using Builders.Commands;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.ViewModels
{
    public class DIC_addItemViewModel : ViewModel
    {
        public string WindowName { get; } = "Tools Items";
        private string nameItem;
        private bool pressButton;
        
        public string NameItem
        {
            get { return nameItem; }
            set
            {
                nameItem = value;
                OnPropertyChanged(nameof(NameItem));
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
            if (NameItem != "")
            {
                if (obj is System.Windows.Window)
                {
                    PressButton = true;
                    (obj as System.Windows.Window).Close();
                }
            }
        }));

        public DIC_addItemViewModel()
        {
            PressButton = false;
        }
    }
}
