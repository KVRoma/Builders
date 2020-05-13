using Builders.Commands;
using Builders.Models;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Builders.ViewModels
{
    public class DIC_SupplierViewModel : ViewModel
    {
        private BuilderContext db;
        public string NameWindow { get; set; } = "Suppliers";

        private DIC_Supplier supplierSelect;
        private IEnumerable<DIC_Supplier> suppliers;
        private DIC_GroupeQuotation groupeSelect;
        private IEnumerable<DIC_GroupeQuotation> groupes;
        private bool pressOk;
        private Visibility visibilityMenu;
        private Visibility visibilitySelect;
        private Visibility visibilityEdit;
        private bool isEnableButton;
        private bool isCreated;

        public DIC_Supplier SupplierSelect
        {
            get { return supplierSelect; }
            set
            {
                supplierSelect = value;
                OnPropertyChanged(nameof(SupplierSelect));
                if (SupplierSelect != null)
                {
                    IsEnableButton = true;
                }
                else
                {
                    IsEnableButton = false;
                }
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
        public DIC_GroupeQuotation GroupeSelect
        {
            get { return groupeSelect; }
            set
            {
                groupeSelect = value;
                OnPropertyChanged(nameof(GroupeSelect));
                if (GroupeSelect != null)
                {
                    Suppliers = db.DIC_Suppliers.Local.ToBindingList().Where(s => s.GroupeId == GroupeSelect.Id).OrderBy(s => s.Supplier);
                }
            }
        }
        public IEnumerable<DIC_GroupeQuotation> Groupes
        {
            get { return groupes; }
            set
            {
                groupes = value;
                OnPropertyChanged(nameof(Groupes));
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
        public Visibility VisibilityMenu
        {
            get { return visibilityMenu; }
            set
            {
                visibilityMenu = value;
                OnPropertyChanged(nameof(VisibilityMenu));
            }
        }
        public Visibility VisibilitySelect
        {
            get { return visibilitySelect; }
            set
            {
                visibilitySelect = value;
                OnPropertyChanged(nameof(VisibilitySelect));
            }
        }
        public Visibility VisibilityEdit
        {
            get { return visibilityEdit; }
            set 
            {
                visibilityEdit = value;
                OnPropertyChanged(nameof(VisibilityEdit));
            }
        }
        public bool IsEnableButton
        {
            get { return isEnableButton; }
            set
            {
                isEnableButton = value;
                OnPropertyChanged(nameof(IsEnableButton));
            }
        }
        public bool IsCreated
        {
            get { return isCreated; }
            set
            {
                isCreated = value;
                OnPropertyChanged(nameof(IsCreated));
            }
        }

        private Command selectToExitCommand;
        private Command addCommand;
        private Command insCommand;
        private Command delCommand;
        private Command okCommand;
        private Command cancelCommand;

        public Command SelectToExitCommand => selectToExitCommand ?? (selectToExitCommand = new Command(obj=> 
        {
            if (obj is System.Windows.Window)
            {
                PressOk = true;
                (obj as System.Windows.Window).Close();
            }
        }));
        public Command AddCommand => addCommand ?? (addCommand = new Command(obj=> 
        {
            IsCreated = true;
            SupplierSelect = new DIC_Supplier();
            VisibilityEdit = Visibility.Visible;
            VisibilityMenu = Visibility.Collapsed;
            
        }));
        public Command InsCommand => insCommand ?? (insCommand = new Command(obj=> 
        {
            IsCreated = false;
            VisibilityEdit = Visibility.Visible;
            VisibilityMenu = Visibility.Collapsed;
        }));
        public Command DelCommand => delCommand ?? (delCommand = new Command(obj=> 
        {
            db.DIC_Suppliers.Remove(SupplierSelect);
            db.SaveChanges();
            Suppliers = db.DIC_Suppliers.Local.ToBindingList().Where(s => s.GroupeId == GroupeSelect.Id).OrderBy(s => s.Supplier);
        }));
        public Command OkCommand => okCommand ?? (okCommand = new Command(obj=> 
        {            
            VisibilityEdit = Visibility.Collapsed;
            VisibilityMenu = Visibility.Visible;
            if (IsCreated)
            {
                db.DIC_Suppliers.Add(SupplierSelect);
                IsCreated = false;
            }
            else
            {
                db.Entry(SupplierSelect).State = EntityState.Modified;
                IsCreated = false;
            }            
            db.SaveChanges();
            Suppliers = db.DIC_Suppliers.Local.ToBindingList().Where(s => s.GroupeId == GroupeSelect.Id).OrderBy(s => s.Supplier);
        }));
        public Command CancelCommand => cancelCommand ?? (cancelCommand = new Command(obj=> 
        {            
            VisibilityEdit = Visibility.Collapsed;
            VisibilityMenu = Visibility.Visible;
        }));

        public DIC_SupplierViewModel(ref BuilderContext context, int? idGroupe)
        {
            VisibilitySelect = Visibility.Visible;
            VisibilityMenu = Visibility.Collapsed;
            VisibilityEdit = Visibility.Collapsed;

            db = context;
            PressOk = false;
            db.DIC_Suppliers.Load();
            Suppliers = db.DIC_Suppliers.Local.ToBindingList().Where(s=>s.GroupeId == idGroupe).OrderBy(s=>s.Supplier);
        }
        public DIC_SupplierViewModel(ref BuilderContext context)
        {
            VisibilityMenu = Visibility.Visible;
            VisibilitySelect = Visibility.Collapsed;
            VisibilityEdit = Visibility.Collapsed;
            IsEnableButton = false;
            IsCreated = false;

            db = context;
            PressOk = false;
            db.DIC_Suppliers.Load();
            db.DIC_GroupeQuotations.Load();
            Groupes = db.DIC_GroupeQuotations.Local.ToBindingList().Where(g=>g.Id == 1 | g.Id == 2);
        }

    }
}
