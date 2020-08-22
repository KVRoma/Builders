using Builders.Commands;
using Builders.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.ViewModels
{
    public class QuotaItemEditViewModel : ViewModel
    {
        private decimal quantity;
        private decimal price;
        private bool pressOk;

        public string WindowName { get; } = "Data user ...";
        public decimal Quantity
        {
            get { return quantity; }
            set
            {
                quantity = value;
                OnPropertyChanged(nameof(Quantity));
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
        public bool PressOk
        {
            get { return pressOk; }
            set
            {
                pressOk = value;
                OnPropertyChanged(nameof(PressOk));
            }
        }

        public QuotaItemEditViewModel()
        {
            PressOk = false;
        }

        private Command okCommand;
        private Command cancelCommand;

        public Command OkCommand => okCommand ?? (okCommand = new Command(obj=> 
        {
            if (obj is System.Windows.Window)
            {
                PressOk = true;
                (obj as System.Windows.Window).Close();
            }
        }));

        public Command CancelCommand => cancelCommand ?? (cancelCommand = new Command(obj=> 
        {
            if (obj is System.Windows.Window)
            {
                PressOk = false;
                (obj as System.Windows.Window).Close();
            }
        }));

       
    }
}
