using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Builders.Commands;
using Builders.Models;

namespace Builders.ViewModels
{
    public class InvoiceViewModel : ViewModel
    {
        public string WindowName { get; } = "Invoice";
        
        private Quotation quotaSelect;
        private IEnumerable<Quotation> quota;
        private string upNumber;
        private string orderNumber;
        private bool pressOk;

        public Quotation QuotaSelect
        {
            get => quotaSelect;
            set 
            { 
                quotaSelect = value;
                OnPropertyChanged(nameof(QuotaSelect));
            }
        }
        public IEnumerable<Quotation> Quota
        {
            get => quota;
            set 
            { 
                quota = value;
                OnPropertyChanged(nameof(Quota));
            }
        }
        public string UpNumber
        {
            get => upNumber;
            set 
            { 
                upNumber = value;
                OnPropertyChanged(nameof(UpNumber));
            }
        }
        public string OrderNumber
        {
            get => orderNumber;
            set 
            { 
                orderNumber = value;
                OnPropertyChanged(nameof(OrderNumber));
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
            if (QuotaSelect != null)
            {
                if (obj is System.Windows.Window)
                {
                    PressOk = true;
                    (obj as System.Windows.Window).Close();
                }
            }
        }));

        public InvoiceViewModel()
        {
            PressOk = false;
        }
    }
}
