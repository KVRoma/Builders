using Builders.Commands;
using Builders.Enums;
using Builders.Models;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.ViewModels
{
    public class DIC_ClientViewModel : ViewModel
    {
        public string WindowName { get; set; } = "Builders - (Dictionarys)";
        private BuilderContext db;

        private IdicClient itemSelect;
        private IEnumerable<IdicClient> items;
        private EnumDictionary result;

        public IdicClient ItemSelect
        {
            get { return itemSelect; }
            set
            {
                itemSelect = value;
                OnPropertyChanged(nameof(ItemSelect));
            }
        }
        public IEnumerable<IdicClient> Items
        {
            get { return items; }
            set
            {
                items = value;
                OnPropertyChanged(nameof(Items));
            }
        }
        public EnumDictionary Result
        {
            get { return result; }
            set
            {
                result = value;
                OnPropertyChanged(nameof(Result));
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
                switch (Result)
                {
                    case EnumDictionary.TypeOfClient:
                        {
                            var typeOfClient = db.DIC_TypeOfClients.Find(ItemSelect.Id);
                            db.DIC_TypeOfClients.Remove(typeOfClient);
                        }
                        break;
                    case EnumDictionary.HearAboutsUs:
                        {
                            var abouts = db.DIC_HearAboutsUse.Find(ItemSelect.Id);
                            db.DIC_HearAboutsUse.Remove(abouts);
                        }
                        break;
                }
                db.SaveChanges();
            }
        }));

        public DIC_ClientViewModel(ref BuilderContext context, EnumDictionary res)
        {
            db = context;
            Result = res;
            switch (Result)
            {
                case EnumDictionary.TypeOfClient:
                    {
                        db.DIC_TypeOfClients.Load();
                        Items = db.DIC_TypeOfClients.Local.ToBindingList();
                    }
                    break;
                case EnumDictionary.HearAboutsUs:
                    {
                        db.DIC_HearAboutsUse.Load();
                        Items = db.DIC_HearAboutsUse.Local.ToBindingList();
                    }
                    break;
            }
        }
    }
}
