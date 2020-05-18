using Builders.Commands;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.ViewModels
{
    public class DeliveryViewModel : ViewModel
    {
        public string NameWindow { get; set; } = "Delivery";
        private string orderNumber;
        private decimal amountDelivery;
        private bool pressOk;

        public string OrderNumber
        {
            get { return orderNumber; }
            set
            {
                orderNumber = value;
                OnPropertyChanged(nameof(OrderNumber));
            }
        }
        public decimal AmountDelivery
        {
            get { return amountDelivery; }
            set
            {
                amountDelivery = value;
                OnPropertyChanged(nameof(AmountDelivery));
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
        public Command OkCommand => okCommand ?? (okCommand = new Command(obj=> 
        {
            if (obj is System.Windows.Window)
            {
                PressOk = true;
                (obj as System.Windows.Window).Close();
            }
        }));

        public DeliveryViewModel()
        {
            PressOk = false;
            AmountDelivery = 0m;
        }
    }
}
