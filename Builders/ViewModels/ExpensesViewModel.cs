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
    public class ExpensesViewModel : ViewModel
    {
        private BuilderContext db;
        public string NameWindow { get; } = "Expenses";

        private IEnumerable<Expenses> expenses;
        private IList<Expenses> expensesSelect;
        private bool enableButtonGrid;
        private Visibility visibleView;
        private Visibility visibleAdd;
        private DateTime date;
        private string typeSelect;
        private List<string> types;
        private string name;
        private string description;
        private decimal amount;
        private bool flagEdit;
        private DateTime localdateFrom;
        private DateTime localdateTo;


        public IEnumerable<Expenses> Expenses
        {
            get { return expenses; }
            set
            {
                expenses = value;
                OnPropertyChanged(nameof(Expenses));
            }
        }
        public IList<Expenses> ExpensesSelect
        {
            get { return expensesSelect; }
            set
            {
                expensesSelect = value;
                OnPropertyChanged(nameof(ExpensesSelect));
                if (ExpensesSelect != null)
                {
                    EnableButtonGrid = true;
                }
                else
                {
                    EnableButtonGrid = false;
                }
            }
        }
        public bool EnableButtonGrid
        {
            get { return enableButtonGrid; }
            set
            {
                enableButtonGrid = value;
                OnPropertyChanged(nameof(EnableButtonGrid));
            }
        }
        public Visibility VisibleView
        {
            get { return visibleView; }
            set
            {
                visibleView = value;
                OnPropertyChanged(nameof(VisibleView));
            }
        }
        public Visibility VisibleAdd
        {
            get {return visibleAdd; }
            set
            {
                visibleAdd = value;
                OnPropertyChanged(nameof(VisibleAdd));
            }
        }
        public DateTime Date
        {
            get { return date; }
            set
            {
                date = value;
                OnPropertyChanged(nameof(Date));
            }
        }
        public string TypeSelect        
        {
            get { return typeSelect; }
            set
            {
                typeSelect = value;
                OnPropertyChanged(nameof(TypeSelect));
            }
        }
        public List<string> Types
        {
            get { return types; }
            set
            {
                types = value;
                OnPropertyChanged(nameof(Types));
            }
        }
        public string Name
        {
            get { return name; }
            set
            {
                name = value;
                OnPropertyChanged(nameof(Name));
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

        private Command _addCommand;
        private Command _insCommand;
        private Command _delCommand;
        private Command _okCommand;


        public Command AddCommand => _addCommand ?? (_addCommand = new Command(obj=> 
        {
            VisibleView = Visibility.Collapsed;
            VisibleAdd = Visibility.Visible;

            flagEdit = false;

            Date = DateTime.Today;
            TypeSelect = Types[0];
            Name = "";
            Description = "";
            Amount = 0m;            
        }));
        public Command InsCommand => _insCommand ?? (_insCommand = new Command(obj=> 
        {
            if (ExpensesSelect.Count() > 0)
            {
                VisibleView = Visibility.Collapsed;
                VisibleAdd = Visibility.Visible;

                flagEdit = true;

                Date = ExpensesSelect[0].Date;
                TypeSelect = ExpensesSelect[0].Type;
                Name = ExpensesSelect[0].Name;
                Description = ExpensesSelect[0].Description;
                Amount = ExpensesSelect[0].Amounts;
            }
        }));
        public Command DelCommand => _delCommand ?? (_delCommand = new Command(obj=> 
        {
            if (ExpensesSelect != null)
            {
               
               db.Expenses.RemoveRange(ExpensesSelect);
               db.SaveChanges();

                Expenses = null;
                Expenses = db.Expenses.Local.ToBindingList().Where(i => i.Date >= localdateFrom && i.Date <= localdateTo).OrderBy(e => e.Type).ThenBy(e => e.Date);
            }
        }));

        public Command OkCommand => _okCommand ?? (_okCommand = new Command(obj=> 
        {           
            Expenses expenses = new Expenses() 
            {
                Date = Date,
                Type = TypeSelect,
                Name = Name,
                Description = Description,
                Amounts = Amount
            };

            if (flagEdit)
            {
                ExpensesSelect[0].Date = expenses.Date;
                ExpensesSelect[0].Type = expenses.Type;
                ExpensesSelect[0].Name = expenses.Name;
                ExpensesSelect[0].Description = expenses.Description;
                ExpensesSelect[0].Amounts = expenses.Amounts;

                db.Entry(ExpensesSelect[0]).State = EntityState.Modified;
                db.SaveChanges();
            }
            else
            {
                db.Expenses.Add(expenses);
                db.SaveChanges();
            }
            

            Expenses = null;
            Expenses = db.Expenses.Local.ToBindingList().Where(i => i.Date >= localdateFrom && i.Date <= localdateTo).OrderBy(e => e.Type).ThenBy(e => e.Date);

            VisibleView = Visibility.Visible;
            VisibleAdd = Visibility.Collapsed;
        }));

        public ExpensesViewModel(DateTime dateFrom, DateTime dateTo, ref BuilderContext context)
        {
            db = context;
            localdateFrom = dateFrom;
            localdateTo = dateTo;
            db.Expenses.Load();            
            Expenses = db.Expenses.Local.ToBindingList().Where(i => i.Date >= localdateFrom && i.Date <= localdateTo).OrderBy(e=>e.Type).ThenBy(e=>e.Date);            
            ExpensesSelect = new List<Expenses>();            
            EnableButtonGrid = false;
            VisibleView = Visibility.Visible;
            VisibleAdd = Visibility.Collapsed;

            Types = new List<string>() 
            {
                "Fixed", 
                "Variable"
            };
            
        }
    }
}
