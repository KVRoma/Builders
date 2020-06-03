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
    public class DIC_DepthViewModel : ViewModel
    {
        public string WindowName { get; set; } = "Depth";
        private BuilderContext db;

        private DIC_DepthQuotation itemSelect;
        private IEnumerable<DIC_DepthQuotation> items;


        public DIC_DepthQuotation ItemSelect
        {
            get { return itemSelect; }
            set
            {
                itemSelect = value;
                OnPropertyChanged(nameof(ItemSelect));
            }
        }
        public IEnumerable<DIC_DepthQuotation> Items
        {
            get { return items; }
            set
            {
                items = value;
                OnPropertyChanged(nameof(Items));
            }
        }


        private Command _saveCommand;
        private Command _delCommand;

        public Command SaveCommand => _saveCommand ?? (_saveCommand = new Command(obj =>
        {
            db.SaveChanges();
            if (obj is System.Windows.Window)
            {
                (obj as System.Windows.Window).Close();
            }
        }));
        public Command DelCommand => _delCommand ?? (_delCommand = new Command(obj =>
        {
            if (ItemSelect != null)
            {               
                db.DIC_DepthQuotations.Remove(ItemSelect);
                db.SaveChanges();
            }
        }));

        public DIC_DepthViewModel(ref BuilderContext context)
        {
            db = context;
            db.DIC_DepthQuotations.Load();
            Items = db.DIC_DepthQuotations.Local.ToBindingList();
        }
    }
}
