using Builders.Commands;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.ViewModels
{
    public class QuotaOtherViewModel : ViewModel
    {
        public string WindowName { get; } = "Quotation - (Other)";
        private string description;
        private string notes;
        private decimal discountMaterial;
        private decimal discountLabour;
        private bool financing;
        private decimal creditCard;
        private bool pressOk;

        public string Description
        {
            get { return description; }
            set
            {
                description = value;
                OnPropertyChanged(nameof(Description));
            }
        }
        public string Notes
        {
            get { return notes; }
            set 
            { 
                notes = value;
                OnPropertyChanged(nameof(Notes));
            } 
        }
        public decimal DiscountMaterial
        {
            get { return discountMaterial; }
            set
            {
                discountMaterial = value;
                OnPropertyChanged(nameof(DiscountMaterial));
            }
        }
        public decimal DiscountLabour
        {
            get { return discountLabour; }
            set
            {
                discountLabour = value;
                OnPropertyChanged(nameof(DiscountLabour));
            }
        }
        public bool Financing 
        {
            get { return financing; }
            set 
            {
                financing = value;
                OnPropertyChanged(nameof(Financing));
            } 
        }
        public decimal CreditCard 
        {
            get {return creditCard; }
            set 
            {
                creditCard = value;
                OnPropertyChanged(nameof(CreditCard));
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

        private Command _okCommand;

        public Command OkCommand => _okCommand ?? (_okCommand = new Command(obj => 
        {            
            if (obj is System.Windows.Window)
            {
                PressOk = true;
                (obj as System.Windows.Window).Close();
            }            
        }));

        public QuotaOtherViewModel()
        {
            PressOk = false;
        }
    }
}
