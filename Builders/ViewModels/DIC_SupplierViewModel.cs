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
    public class DIC_SupplierViewModel : ViewModel
    {
        private BuilderContext db;
        public string NameWindow { get; set; } = "Suppliers";

        private DIC_Supplier supplierSelect;
        private IEnumerable<DIC_Supplier> suppliers;
        private bool pressOk;

        public DIC_Supplier SupplierSelect
        {
            get { return supplierSelect; }
            set
            {
                supplierSelect = value;
                OnPropertyChanged(nameof(SupplierSelect));
            }
        }
        public IEnumerable<DIC_Supplier> Suppliers
        {
            get { return suppliers; }
            set
            {
                suppliers = value;
                OnPropertyChanged(nameof(Suppliers));
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

        private Command selectToExitCommand;
        public Command SelectToExitCommand => selectToExitCommand ?? (selectToExitCommand = new Command(obj=> 
        {
            if (obj is System.Windows.Window)
            {
                PressOk = true;
                (obj as System.Windows.Window).Close();
            }
        }));

        public DIC_SupplierViewModel(ref BuilderContext context, int? idGroupe)
        {
            db = context;
            PressOk = false;
            db.DIC_Suppliers.Load();
            Suppliers = db.DIC_Suppliers.Local.ToBindingList().Where(s=>s.GroupeId == idGroupe).OrderBy(s=>s.Supplier);
        }
        public DIC_SupplierViewModel(ref BuilderContext context)
        {
            db = context;
            PressOk = false;
            db.DIC_Suppliers.Load();            
        }

    }
}
