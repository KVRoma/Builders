using Builders.Commands;
using Builders.Enums;
using Builders.Models;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace Builders.ViewModels
{
    public class DIC_QuotationViewModel : ViewModel
    {
        public string WindowName { get; } = "Builders - (Dictionary)";
        private BuilderContext db;

        private DIC_GroupeQuotation groupeSelect;
        private IEnumerable<DIC_GroupeQuotation> groupes;
        private DIC_ItemQuotation itemSelect;
        private IEnumerable<DIC_ItemQuotation> items;
        private DIC_DescriptionQuotation descriptionSelect;
        private IEnumerable<DIC_DescriptionQuotation> descriptions;
        

        private EnumDictionary result;
        
        public DIC_GroupeQuotation GroupeSelect
        {
            get { return groupeSelect; }
            set
            {
                groupeSelect = value;
                OnPropertyChanged(nameof(GroupeSelect));                
                Items = (GroupeSelect != null) ? (db.DIC_ItemQuotations.Local.ToBindingList().Where(t => t.GroupeId == GroupeSelect.Id).OrderBy(t=>t.Name)) : null;                
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
        public DIC_ItemQuotation ItemSelect
        {
            get { return itemSelect; }
            set
            {
                itemSelect = value;
                OnPropertyChanged(nameof(ItemSelect));
                Descriptions = (ItemSelect != null) ? (db.DIC_DescriptionQuotations.Local.ToBindingList().Where(t => t.ItemId == ItemSelect.Id).OrderBy(t => t.Name)) : null;
            }
        }
        public IEnumerable<DIC_ItemQuotation> Items
        {
            get { return items; }
            set
            {
                items = value;
                OnPropertyChanged(nameof(Items));
            }
        }
        public DIC_DescriptionQuotation DescriptionSelect
        {
            get { return descriptionSelect; }
            set
            {
                descriptionSelect = value;
                OnPropertyChanged(nameof(DescriptionSelect));
            }
        }
        public IEnumerable<DIC_DescriptionQuotation> Descriptions
        {
            get { return descriptions; }
            set
            {
                descriptions = value;
                OnPropertyChanged(nameof(Descriptions));
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
       

        private Command _addItemCommand;
        private Command _insItemCommand;
        private Command _delItemCommand;
        private Command _dubleClickCommand;
        //**************************************************
        private Command _addDescriptionCommand;
        private Command _insDescriptionCommand;
        private Command _delDescriptionCommand;
        private Command _loadDescriptionExcel;

        private Command _search;

        public Command AddItemCommand => _addItemCommand ?? (_addItemCommand = new Command(async obj =>
        {
            var displayRootRegistry = (Application.Current as App).displayRootRegistry;
            var addItemViewModel = new DIC_addItemViewModel() {NameItem = "" };
            await displayRootRegistry.ShowModalPresentation(addItemViewModel);
            if (addItemViewModel.PressButton)
            {
                DIC_ItemQuotation itemNew = new DIC_ItemQuotation()
                {
                    Name = addItemViewModel.NameItem,
                    GroupeId = GroupeSelect.Id,
                };
                db.DIC_ItemQuotations.Add(itemNew);
                db.SaveChanges();
                Items = null;
                Items = (GroupeSelect != null) ? (db.DIC_ItemQuotations.Local.ToBindingList().Where(t => t.GroupeId == GroupeSelect.Id)) : null;
            }
        }));
        public Command InsItemCommand => _insItemCommand ?? (_insItemCommand = new Command(async obj =>
        {
            if (ItemSelect != null)
            {
                var displayRootRegistry = (Application.Current as App).displayRootRegistry;
                var insItemViewModel = new DIC_addItemViewModel() { NameItem = ItemSelect.Name };
                await displayRootRegistry.ShowModalPresentation(insItemViewModel);
                if (insItemViewModel.PressButton)
                {
                    ItemSelect.Name = insItemViewModel.NameItem;
                    db.Entry(ItemSelect).State = EntityState.Modified;
                    db.SaveChanges();
                    Items = null;
                    Items = (GroupeSelect != null) ? (db.DIC_ItemQuotations.Local.ToBindingList().Where(t => t.GroupeId == GroupeSelect.Id)) : null;
                }
            }
        }));
        public Command DelItemCommand => _delItemCommand ?? (_delItemCommand = new Command(obj =>
        {
            if (ItemSelect != null)
            {
                var des = db.DIC_DescriptionQuotations.Where(d => d.ItemId == ItemSelect.Id);
                db.DIC_DescriptionQuotations.RemoveRange(des);
                db.DIC_ItemQuotations.Remove(ItemSelect);
                db.SaveChanges();
                Items = null;
                Items = (GroupeSelect != null) ? (db.DIC_ItemQuotations.Local.ToBindingList().Where(t => t.GroupeId == GroupeSelect.Id)) : null;
            }           
        }));
        public Command DubleClickCommand => _dubleClickCommand ?? (_dubleClickCommand = new Command(async obj=> 
        {
            if (GroupeSelect.Id == 1 || GroupeSelect.Id == 2)
            {
                var displayRootRegistry = (Application.Current as App).displayRootRegistry;
                var supplierViewModel = new DIC_SupplierViewModel(ref db, GroupeSelect.Id);
                await displayRootRegistry.ShowModalPresentation(supplierViewModel);
                if (supplierViewModel.PressOk)
                {
                    ItemSelect.SupplierId = supplierViewModel.SupplierSelect.Id;
                    ItemSelect.Color = "Blue";
                    db.Entry(ItemSelect).State = EntityState.Modified;
                    db.SaveChanges();
                    Items = db.DIC_ItemQuotations.Local.ToBindingList().Where(t => t.GroupeId == GroupeSelect.Id).OrderBy(t => t.Name);
                }
            }
        }));
        //*************************************************
        public Command AddDescriptionCommand => _addDescriptionCommand ?? (_addDescriptionCommand = new Command(async obj => 
        {
            var displayRootRegistry = (Application.Current as App).displayRootRegistry;
            var addDesViewModel = new DIC_addDescriptionViewModel() {NameDescription = "", Price = 0, Cost = 0 };
            await displayRootRegistry.ShowModalPresentation(addDesViewModel);
            if (addDesViewModel.PressButton)
            {
                DIC_DescriptionQuotation descNew = new DIC_DescriptionQuotation()
                {
                    Name = addDesViewModel.NameDescription,
                    Price = addDesViewModel.Price,
                    Cost = addDesViewModel.Cost,
                    ItemId = ItemSelect.Id
                };
                db.DIC_DescriptionQuotations.Add(descNew);
                db.SaveChanges();
                Descriptions = null;
                Descriptions = (ItemSelect != null) ? (db.DIC_DescriptionQuotations.Local.ToBindingList().Where(t => t.ItemId == ItemSelect.Id)) : null;
            }
        }));
        public Command InsDescriptionCommand => _insDescriptionCommand ?? (_insDescriptionCommand = new Command(async obj =>
        {
            if (DescriptionSelect != null)
            {
                var displayRootRegistry = (Application.Current as App).displayRootRegistry;
                var insDesViewModel = new DIC_addDescriptionViewModel() 
                {
                    NameDescription = DescriptionSelect.Name, 
                    Price = DescriptionSelect.Price, 
                    Cost = DescriptionSelect.Cost 
                };
                await displayRootRegistry.ShowModalPresentation(insDesViewModel);
                if (insDesViewModel.PressButton)
                {
                    DescriptionSelect.Name = insDesViewModel.NameDescription;
                    DescriptionSelect.Price = insDesViewModel.Price;
                    DescriptionSelect.Cost = insDesViewModel.Cost;
                    db.Entry(DescriptionSelect).State = EntityState.Modified;
                    db.SaveChanges();
                    Descriptions = null;
                    Descriptions = (ItemSelect != null) ? (db.DIC_DescriptionQuotations.Local.ToBindingList().Where(t => t.ItemId == ItemSelect.Id)) : null;
                }
            }
        }));
        public Command DelDescriptionCommand => _delDescriptionCommand ?? (_delDescriptionCommand = new Command(obj =>
        {
            db.DIC_DescriptionQuotations.Remove(DescriptionSelect);
            db.SaveChanges();
            Descriptions = null;
            Descriptions = (ItemSelect != null) ? (db.DIC_DescriptionQuotations.Local.ToBindingList().Where(t => t.ItemId == ItemSelect.Id)) : null;
        }));
        public Command LoadDescriptionExcel => _loadDescriptionExcel ?? (_loadDescriptionExcel = new Command(obj=> 
        {
            ImportDescriptionToExcel();
        }));

        public Command Search => _search ?? (_search = new Command(obj=> 
        {
            string search = obj.ToString();
            if (search == "")
            {
                Items = (GroupeSelect != null) ? (db.DIC_ItemQuotations.Local.ToBindingList().Where(t => t.GroupeId == GroupeSelect.Id).OrderBy(t => t.Name)) : null;
            }
            else
            {                
                Items = (GroupeSelect != null) ? (Items.Where(t => t.GroupeId == GroupeSelect.Id)
                                                       .Where(t => t.Name.ToUpper().Contains(search.ToUpper()))
                                                       .OrderBy(t => t.Name)) : null;
            }
        }));

        public DIC_QuotationViewModel(ref BuilderContext context, EnumDictionary res)
        {
            db = context;
            Result = res;
            db.DIC_GroupeQuotations.Load();
            db.DIC_ItemQuotations.Load();
            db.DIC_DescriptionQuotations.Load();
            Groupes = db.DIC_GroupeQuotations.Local.ToBindingList();
        }

        /// <summary>
        /// Відкриває файл по заданій масці (один файл)
        /// </summary>
        /// <returns></returns>
        private string OpenFile(string filter)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = filter;
            if (dialog.ShowDialog() == true)
            {
                return dialog.FileName;
            }
            return null;
        }
        /// <summary>
        /// Завантажує назви та ціни з файлу
        /// </summary>
        private async void ImportDescriptionToExcel()
        {
            string path = OpenFile("Файл Excel|*.XLSX;*.XLS;*.XLSM");   // Вибираємо наш файл (метод OpenFile() описаний нижче)

            if (path == null) // Перевіряємо шлях до файлу на null
            {
                return;
            }
            List<DIC_DescriptionQuotation> descriptionQuotations = new List<DIC_DescriptionQuotation>();
            
            await Start(path);

            db.DIC_DescriptionQuotations.AddRange(descriptionQuotations);
            db.SaveChanges();

            var temp = ItemSelect;
            ItemSelect = temp;


            async Task Start(string pathFile)
            {
                await Task.Run(() =>
                {
                    Excel.Application ExcelApp = new Excel.Application();     // Створюємо додаток Excel
                    Excel.Workbook ExcelWorkBook;                             // Створюємо книгу Excel
                    Excel.Worksheet ExcelWorkSheet;                          // Створюємо лист Excel                        

                    try
                    {
                        ExcelWorkBook = ExcelApp.Workbooks.Open(pathFile);                  // Відкриваємо файл Excel                
                        ExcelWorkSheet = ExcelWorkBook.ActiveSheet;                     // Відкриваємо активний Лист Excel                

                        var descrip = Descriptions.ToList();
                        if (descrip != null)
                        {
                            db.DIC_DescriptionQuotations.RemoveRange(descrip);
                            db.SaveChanges();
                        }
                        //int count = int.Parse(ExcelApp.Cells[1, 1].Value2?.ToString()) + 2;
                        int i = 1;
                        while (ExcelApp.Cells[i, 1].Value?.ToString() != null)
                        {
                            descriptionQuotations.Add(new DIC_DescriptionQuotation()
                            {
                                ItemId = ItemSelect.Id,
                                Name = ExcelApp.Cells[i, 1].Value?.ToString(),
                                Price = decimal.TryParse(ExcelApp.Cells[i, 2].Value?.ToString(), out decimal result) ? (result) : (0m),
                                Cost = 0m                                
                            });
                            i++;
                        }
                        ExcelApp.Visible = true;
                        ExcelApp.UserControl = true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                        ExcelApp.Visible = true;
                        ExcelApp.UserControl = true;
                    }
                });
            }
        }
    }
}
