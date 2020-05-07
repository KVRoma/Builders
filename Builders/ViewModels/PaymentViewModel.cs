using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Builders.Commands;
using Builders.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Builders.ViewModels
{
    public class PaymentViewModel : ViewModel
    {
        private BuilderContext db;
        public string WindowName { get; } = "Payment";

        private bool enableDelAndPrint;
        private bool enableAdd;
        
        private Quotation quota;

        private Reciept recieptSelect;
        private IEnumerable<Reciept> reciepts;

        private Payment paymentSelect;
        private IEnumerable<Payment> payments;
        
        private DateTime datePayment;
        private decimal amountPayment;
        private decimal balance;
        private DIC_PaymentMethod methodSelect;
        private IEnumerable<DIC_PaymentMethod> medhods;

        
        public bool EnableDelAndPrint
        {
            get { return enableDelAndPrint; }
            set
            {
                enableDelAndPrint = value;
                OnPropertyChanged(nameof(EnableDelAndPrint));
            }
        }
        public bool EnableAdd
        {
            get { return enableAdd; }
            set
            {
                enableAdd = value;
                OnPropertyChanged(nameof(EnableAdd));
            }
        }
       
        public Reciept RecieptSelect
        {
            get { return recieptSelect; }
            set
            {
                recieptSelect = value;
                OnPropertyChanged(nameof(RecieptSelect));                
            }
        }
        public IEnumerable<Reciept> Reciepts
        {
            get { return reciepts; }
            set
            {
                reciepts = value;
                OnPropertyChanged(nameof(Reciepts));
            }
        }

        public Payment PaymentSelect
        {
            get { return paymentSelect; }
            set
            {
                paymentSelect = value;
                OnPropertyChanged(nameof(PaymentSelect));
                if (PaymentSelect != null)
                {
                    RecieptSelect = Reciepts.FirstOrDefault(r => r.PayNumber == PaymentSelect.Id);
                    EnableDelAndPrint = true;
                }
                else
                {
                    EnableDelAndPrint = false;
                }
            }
        }
        public IEnumerable<Payment> Payments
        {
            get { return payments; }
            set
            {
                payments = value;
                OnPropertyChanged(nameof(Payments));
            }
        }

        public DateTime DatePayment
        {
            get => datePayment;
            set 
            { 
                datePayment = value;
                OnPropertyChanged(nameof(DatePayment));
            }
        }
        public decimal AmountPayment
        {
            get => amountPayment;
            set 
            { 
                amountPayment = value;
                OnPropertyChanged(nameof(AmountPayment));
                if (AmountPayment != 0)
                {
                    DatePayment = (DatePayment == DateTime.MinValue) ? (DateTime.Today) : (DatePayment);
                    CountPay(Payments.Count());
                }
                else
                {
                    DatePayment = DateTime.MinValue;
                    MethodSelect = null;
                    EnableAdd = false;
                }
            }
        }
        public decimal Balance
        {
            get 
            {
                return balance; 
            }
            set 
            { 
                balance = value;
                OnPropertyChanged(nameof(Balance));
            }
        }
        public DIC_PaymentMethod MethodSelect
        {
            get => methodSelect;
            set
            {
                methodSelect = value;
                OnPropertyChanged(nameof(MethodSelect));
            }               
        }
        public IEnumerable<DIC_PaymentMethod> Medhods
        {
            get => medhods;
            set 
            { 
                medhods = value;
                OnPropertyChanged(nameof(Medhods));
            }
        }

              

        private Command addCommand;
        private Command delCommand;
        private Command printCommand;

        public Command AddCommand => addCommand ?? (addCommand = new Command(obj => 
        {
            if (AmountPayment != 0m && MethodSelect != null)
            {
                decimal sum = (Payments.Count() != 0) ? (Payments.Select(p => p.PaymentAmountPaid).Sum()) : (0m);
                decimal balance = decimal.Round(quota.InvoiceGrandTotal - sum - AmountPayment, 2);
                Payment pay = new Payment()
                {
                    PaymentDatePaid = DatePayment,
                    PaymentAmountPaid = AmountPayment,
                    PaymentMethod = MethodSelect.Name,
                    Balance = balance,
                    QuotationId = quota.Id
                };
                db.Payments.Add(pay);
                db.SaveChanges();
                Payments = null;
                Payments = db.Payments.Local.ToBindingList().Where(p => p.QuotationId == quota.Id);
                PaymentSelect = Payments.OrderByDescending(i => i.Id).FirstOrDefault();

                Reciept reciept = new Reciept() 
                {
                    NumberQuota = quota.NumberQuota,
                    PayNumber = PaymentSelect.Id,
                    QuotaId = quota.Id
                };
                db.Reciepts.Add(reciept);
                db.SaveChanges();
                Reciepts = null;
                Reciepts = db.Reciepts.Local.ToBindingList().Where(r => r.QuotaId == quota.Id);
                CountPay(Payments.Count());

                if (balance <= 0)
                {
                    quota.SortingQuota = 2;
                    quota.ActivQuota = true;
                    quota.PaidQuota = true;
                    quota.Color = "Silver";
                }
                else
                {
                    quota.SortingQuota = 1;
                    quota.ActivQuota = true;
                    quota.PaidQuota = false;
                    quota.Color = "Blue";
                }
                db.Entry(quota).State = EntityState.Modified;
                db.SaveChanges();
            }

        }));
        public Command DelCommand => delCommand ?? (delCommand = new Command(obj => 
        {
            if (PaymentSelect != null)
            {
                var pay = Payments.Where(p => p.Id >= PaymentSelect.Id);
                foreach (var item in pay)
                {
                    var reciept = Reciepts.FirstOrDefault(r => r.PayNumber == item.Id);
                    db.Reciepts.Remove(reciept);
                    db.SaveChanges();
                }
                db.Payments.RemoveRange(pay);
                db.SaveChanges();

                Payments = null;
                Payments = db.Payments.Local.ToBindingList().Where(p => p.QuotationId == quota.Id);
                PaymentSelect = Payments.OrderByDescending(i => i.Id).FirstOrDefault();
                
                Reciepts = null;
                Reciepts = db.Reciepts.Local.ToBindingList().Where(r => r.QuotaId == quota.Id);

                CountPay(Payments.Count());

                if (PaymentSelect != null)
                {
                    if (PaymentSelect.Balance <= 0)
                    {
                        quota.SortingQuota = 2;
                        quota.ActivQuota = true;
                        quota.PaidQuota = true;
                        quota.Color = "Silver";
                    }
                    else
                    {
                        quota.SortingQuota = 1;
                        quota.ActivQuota = true;
                        quota.PaidQuota = false;
                        quota.Color = "Blue";
                    }
                }
                else
                {
                    quota.SortingQuota = 3;
                    quota.ActivQuota = false;
                    quota.PaidQuota = false;
                    quota.Color = "Black";
                }
            }
        }));
        public Command PrintCommand => printCommand ?? (printCommand = new Command(obj=> 
        {
            if (PaymentSelect != null)
            {
                Excel.Application ExcelApp = new Excel.Application();
                Excel.Workbook ExcelWorkBook;
                ExcelWorkBook = ExcelApp.Workbooks.Open(Environment.CurrentDirectory + "\\Blanks\\RecieptPDF.xltm");   //Вказуємо шлях до шаблону

                var client = db.Clients.FirstOrDefault(c => c.Id == quota.ClientId);
                var payDo = Payments.Where(p => p.Id <= PaymentSelect.Id).OrderBy(i=>i.Id);                
                
                decimal payCredit = 0m;
                decimal payFee = 0m;
                
                int count = 45;
                foreach (var item in payDo)
                {
                    if (item.PaymentMethod == "Credit Card")
                    {
                        payCredit = decimal.Round(item.PaymentAmountPaid * 1.030927835m, 2);
                        payFee = decimal.Round(payCredit - item.PaymentAmountPaid, 2);
                    }
                    else if (item.PaymentMethod == "Debit Card")
                    {
                        payFee = 0.25m;
                        payCredit = decimal.Round(item.PaymentAmountPaid + payFee, 2);
                    }
                    else
                    {
                        payFee = 0m;
                        payCredit = item.PaymentAmountPaid + payFee;
                    }

                    ExcelApp.Cells[count, 1] = item.PaymentDatePaid;
                    ExcelApp.Cells[count, 2] = payCredit;
                    ExcelApp.Cells[count, 3] = item.PaymentAmountPaid;
                    ExcelApp.Cells[count, 4] = payFee;
                    ExcelApp.Cells[count, 5] = item.PaymentMethod;
                    ExcelApp.Cells[count, 6] = item.Balance;

                    count++;
                }

                if (PaymentSelect.PaymentMethod == "Credit Card")
                {
                    payCredit = decimal.Round(PaymentSelect.PaymentAmountPaid * 1.030927835m, 2);
                    payFee = decimal.Round(payCredit - PaymentSelect.PaymentAmountPaid, 2);
                }
                else if (PaymentSelect.PaymentMethod == "Debit Card")
                {
                    payFee = 0.25m;
                    payCredit = decimal.Round(PaymentSelect.PaymentAmountPaid + payFee, 2);
                }
                else
                {
                    payFee = 0m;
                    payCredit = PaymentSelect.PaymentAmountPaid + payFee;
                }


                ExcelApp.Cells[3, 6] = PaymentSelect.PaymentDatePaid;

                ExcelApp.Cells[38, 6] = payCredit;
                ExcelApp.Cells[39, 6] = PaymentSelect.PaymentMethod;
                ExcelApp.Cells[40, 6] = payFee;
                ExcelApp.Cells[41, 6] = PaymentSelect.PaymentAmountPaid;
                ExcelApp.Cells[34, 6] = payDo.Select(p => p.PaymentAmountPaid).Sum(); // сплачено до цього часу
                            

                ExcelApp.Cells[4, 6] = RecieptSelect?.Number;
                ExcelApp.Cells[6, 6] = quota.NumberQuota;



                ExcelApp.Cells[9, 2] = quota.JobDescription;
                ExcelApp.Cells[10, 2] = client.CompanyName;
                ExcelApp.Cells[11, 2] = client.PrimaryFirstName + " " + client.PrimaryLastName;
                ExcelApp.Cells[12, 2] = client.PrimaryPhoneNumber;
                ExcelApp.Cells[13, 2] = client.PrimaryEmail;
                ExcelApp.Cells[14, 2] = client.AddressBillStreet + ", " + client.AddressBillCity + ", " + client.AddressBillProvince + ", " + client.AddressBillPostalCode + ", " + client.AddressBillCountry;

                ExcelApp.Cells[11, 5] = client.SecondaryFirstName + " " + client.SecondaryLastName;
                ExcelApp.Cells[12, 5] = client.SecondaryPhoneNumber;
                ExcelApp.Cells[13, 5] = client.SecondaryEmail;
                ExcelApp.Cells[14, 5] = client.AddressSiteStreet + ", " + client.AddressSiteCity + ", " + client.AddressSiteProvince + ", " + client.AddressSitePostalCode + ", " + client.AddressSiteCountry;

                ExcelApp.Cells[18, 2] = quota.JobNote;

                ExcelApp.Cells[21, 6] = quota.MaterialSubtotal;
                ExcelApp.Cells[22, 6] = quota.MaterialTax;
                ExcelApp.Cells[23, 6] = quota.MaterialDiscountAmount;
                ExcelApp.Cells[24, 6] = quota.MaterialTotal;

                ExcelApp.Cells[26, 6] = quota.LabourSubtotal;
                ExcelApp.Cells[27, 6] = quota.LabourTax;
                ExcelApp.Cells[28, 6] = quota.LabourDiscountAmount;
                ExcelApp.Cells[29, 6] = quota.LabourTotal;

                ExcelApp.Cells[21, 6] = quota.MaterialSubtotal;
                ExcelApp.Cells[21, 6] = quota.MaterialSubtotal;
                ExcelApp.Cells[21, 6] = quota.MaterialSubtotal;
                ExcelApp.Cells[21, 6] = quota.MaterialSubtotal;
                ExcelApp.Cells[21, 6] = quota.MaterialSubtotal;
                ExcelApp.Cells[21, 6] = quota.MaterialSubtotal;
                ExcelApp.Cells[21, 6] = quota.MaterialSubtotal;
                ExcelApp.Cells[21, 6] = quota.MaterialSubtotal;
                ExcelApp.Cells[21, 6] = quota.MaterialSubtotal;
                ExcelApp.Cells[21, 6] = quota.MaterialSubtotal;

                ExcelApp.Calculate();
                ExcelApp.Cells[1, 9] = "1";   // Записуємо дані в .pdf    
                ExcelApp.Calculate();
            }
        }));

        public PaymentViewModel( ref BuilderContext context, Quotation select)
        {
            db = context;
            quota = select;
            db.DIC_PaymentMethods.Load();
            db.Reciepts.Load();
            db.Payments.Load();

            Medhods = db.DIC_PaymentMethods.Local.ToBindingList();            
            Reciepts = db.Reciepts.Local.ToBindingList().Where(r => r.QuotaId == quota.Id);
            Payments = db.Payments.Local.ToBindingList().Where(p => p.QuotationId == quota.Id);

            EnableDelAndPrint = false;
            EnableAdd = false;
        }
        private void CountPay(int count)
        {
            if (count > 5)
            {
                EnableAdd = false;
            }
            else
            {
                EnableAdd = true;
            }
        }
    }
}
