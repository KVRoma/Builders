using Builders.Commands;
using Builders.Models;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.ViewModels
{
    public class DebtsViewModel : ViewModel
    {
        public string WindowName { get; } = "Debts";
        private Invoice invoiceSelect;
        private IEnumerable<Invoice> invoices;
        private string nameDebts;
        private string description;
        private decimal amount;
        private bool pressOk;
        private bool enableOk;

        public Invoice InvoiceSelect        
        {
            get { return invoiceSelect; }
            set
            {
                invoiceSelect = value;
                OnPropertyChanged(nameof(InvoiceSelect));
                if (InvoiceSelect != null)
                {
                    EnableOk = true;
                }
                else
                {
                    EnableOk = false;
                }
            }
        }
        public IEnumerable<Invoice> Invoices
        {
            get { return invoices; }
            set
            {
                invoices = value;
                OnPropertyChanged(nameof(Invoices));
            }
        }
        public string NameDebts
        {
            get { return nameDebts; }
            set
            {
                nameDebts = value;
                OnPropertyChanged(nameof(NameDebts));
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
        public bool EnableOk
        {
            get { return enableOk; }
            set
            {
                enableOk = value;
                OnPropertyChanged(nameof(EnableOk));
            }
        }

        private Command okCommand;

        public Command OkCommand => okCommand ?? (okCommand = new Command(obj=> 
        {
            if (InvoiceSelect != null)
            {
                if (obj is System.Windows.Window)
                {
                    PressOk = true;
                    (obj as System.Windows.Window).Close();
                }
            }
        }));

        public DebtsViewModel(BuilderContext db, Debts _debtSelect)
        {
            PressOk = false;
            EnableOk = false;

            db.Invoices.Load();
            Invoices = db.Invoices.Local.ToBindingList().Where(i => i.Id > 0);
            
            if (_debtSelect != null)
            {
                InvoiceSelect = db.Invoices.FirstOrDefault(i => i.Id == _debtSelect.InvoiceId);               
            }
            
        }
    }
}
