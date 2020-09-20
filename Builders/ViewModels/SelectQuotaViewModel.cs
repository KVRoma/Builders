using Builders.Commands;
using Builders.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.ViewModels
{
    public class SelectQuotaViewModel : ViewModel
    {
        private bool pressButton;
        private Quotation quotaSelect;
        private IEnumerable<Quotation> quota;
        public string WindowName { get; } = "Quota Select....";
        public bool PressButton
        {
            get { return pressButton; }
            set
            {
                pressButton = value;
            }
        }
        public Quotation QuotaSelect
        {
            get { return quotaSelect; }
            set
            {
                quotaSelect = value;
            }
        }
        public IEnumerable<Quotation> Quota
        {
            get { return quota; }
            set
            {
                quota = value;
            }
        }


        private Command okCommand;

        public Command OkCommand => okCommand ?? (okCommand = new Command(obj =>
        {
            if (obj is System.Windows.Window)
            {
                PressButton = true;
                (obj as System.Windows.Window).Close();
            }
        }));

       
    }
}
