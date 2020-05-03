using Builders.Commands;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.ViewModels
{
    public class DebtsPaymentViewModel : ViewModel
    {
        public string WindowName { get; } = "Payment";
        
        private DateTime date;
        private string description;
        private decimal amount;
        private bool pressOk;
       

        
        public DateTime Date
        {
            get { return date; }
            set
            {
                date = value;
                OnPropertyChanged(nameof(Date));
            }
        }
        public string Description
        {
            get { return description; }
            set
            {
                description = value;
                OnPropertyChanged(nameof(Description));
            }
        }
        public decimal Amount
        {
            get { return amount; }
            set
            {
                amount = value;
                OnPropertyChanged(nameof(Amount));
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
       

        private Command okCommand;

        public Command OkCommand => okCommand ?? (okCommand = new Command(obj =>
        {           
            
            if (obj is System.Windows.Window)
            {
                PressOk = true;
                (obj as System.Windows.Window).Close();
            }
            
        }));

        public DebtsPaymentViewModel()
        {
            PressOk = false;           
            Date = DateTime.Today;
        }
    }
}
