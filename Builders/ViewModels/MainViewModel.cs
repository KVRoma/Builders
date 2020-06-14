using Builders.Models;
using Builders.Commands;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Builders.Enums;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System.IO;
using System.ComponentModel;
using System.Threading;
using Builders.Views;
using System.Windows.Threading;

namespace Builders.ViewModels
{
    public class MainViewModel : ViewModel
    {
        BuilderContext db;
        BackgroundWorker worker;

        public string NameWindow { get; } = "Builder - 2020";
        private Visibility progress;
        private decimal opacityProgress;

        #region Private Property
        private Client clientSelect;
        private IEnumerable<Client> clients;
        private Quotation quotationSelect;
        private IEnumerable<Quotation> quotations;
        private string nameQuotaSelect;
        private List<string> nameQuota;
        private Invoice invoiceSelect;
        private IEnumerable<Invoice> invoices;
        private WorkOrder workOrderSelect;
        private IEnumerable<WorkOrder> workOrders;
        private MaterialProfit materialProfitSelect;
        private IEnumerable<MaterialProfit> materialProfits;
        private LabourProfit labourProfitSelect;
        private IEnumerable<LabourProfit> labourProfits;
        private Debts debtSelect;
        private IEnumerable<Debts> debts;
        private Delivery deliverySelect;
        private IEnumerable<Delivery> deliveries;
        private List<string> deliveriesComboBox;
        private string deliveryComboBoxSelect;
        private bool isCheckedArchiveDelivery;

        private bool enableClient;
        private bool enableQuota;
        private bool enableInvoice;
        private bool enableWorkOrder;
        private bool enableMaterialProfit;
        private bool enableLabourProfit;
        private bool enableDebts;

        private bool isCheckedCMO;
        private bool isCheckedLeveling;
        private Visibility isVisibleCMO;
        private Visibility isVisibleLeveling;
        private string companyName;

        private bool isCheckedProfit;
        private bool isCheckedTotal;
        private bool isCheckedPayroll;
        private bool isCheckedAmount;
        private bool isCheckedDebts;
        private bool isCheckedExpenses;
        private bool enableReport;
        private Visibility isVisibleMenuReport;
        private Visibility isVisibleProfitReport;
        private Visibility isVisibleTotalReport;
        private Visibility isVisiblePayrollReport;
        private Visibility isVisibleAmountReport;
        private Visibility isVisibleDebtsReport;
        private Visibility isVisibleExpensesReport;
        private Visibility isVisibleDateSelect;
        private Visibility isVisibleDateSelectExpenses;
        private DateTime reportDateFrom;
        private DateTime reportDateTo;
        private MaterialProfit materialReport;
        private LabourProfit labourReport;
        private List<Report> totalReport;
        private List<ReportPayroll> reportPayrolls;
        private List<ReportPayrollToWork> reportPayrollToWorks;
        private List<ReportAmount> reportAmounts;
        private ReportTotal totalReportForLabel;
        private List<Debts> amountDebtsReport;
        private List<Debts> paymentDebtsReport;
        private List<Expenses> activeExpenses;
        private Expenses activeExpenseSelect;
        private List<Expenses> paymentExpenses;
        private Expenses paymentExpenseSelect;
        private string countReport;
        private decimal labelAmountDebts;
        private decimal labelPaymentDebts;
        private decimal labelActiveExpenses;
        private decimal labelPaymentExpenses;
        private int dateYearSelect;
        private List<int> dateYears;
        private EnumMonths dateMonthSelect;
        private List<EnumMonths> dateMonths;
        #endregion
        #region Public Property
        public Visibility Progress
        {
            get { return progress; }
            set
            {
                progress = value;
                OnPropertyChanged(nameof(Progress));
            }
        }
        public decimal OpacityProgress
        {
            get { return opacityProgress; }
            set
            {
                opacityProgress = value;
                OnPropertyChanged(nameof(OpacityProgress));
            }
        }

        public Client ClientSelect
        {
            get { return clientSelect; }
            set
            {
                clientSelect = value;
                OnPropertyChanged(nameof(ClientSelect));
                if (ClientSelect == null)
                {
                    EnableClient = false;
                }
                else
                {
                    EnableClient = true;
                }
            }
        }
        public IEnumerable<Client> Clients
        {
            get { return clients; }
            set
            {
                clients = value;
                OnPropertyChanged(nameof(Clients));
            }
        }
        public Quotation QuotationSelect
        {
            get { return quotationSelect; }
            set
            {
                quotationSelect = value;
                OnPropertyChanged(nameof(QuotationSelect));
                if (QuotationSelect == null)
                {
                    EnableQuota = false;
                }
                else
                {
                    EnableQuota = true;
                }
            }
        }
        public IEnumerable<Quotation> Quotations
        {
            get { return quotations; }
            set
            {
                quotations = value;
                OnPropertyChanged(nameof(Quotations));
            }
        }
        public string NameQuotaSelect
        {
            get { return nameQuotaSelect; }
            set
            {
                nameQuotaSelect = value;
                OnPropertyChanged(nameof(NameQuotaSelect));
            }
        }
        public List<string> NameQuota
        {
            get { return nameQuota; }
            set
            {
                nameQuota = value;
                OnPropertyChanged(nameof(NameQuota));
            }
        }

        public Invoice InvoiceSelect
        {
            get { return invoiceSelect; }
            set
            {
                invoiceSelect = value;
                OnPropertyChanged(nameof(InvoiceSelect));
                if (InvoiceSelect == null)
                {
                    EnableInvoice = false;
                }
                else
                {
                    EnableInvoice = true;
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
        public WorkOrder WorkOrderSelect
        {
            get { return workOrderSelect; }
            set
            {
                workOrderSelect = value;
                OnPropertyChanged(nameof(WorkOrderSelect));
                if (WorkOrderSelect == null)
                {
                    EnableWorkOrder = false;
                }
                else
                {
                    EnableWorkOrder = true;
                }
            }
        }
        public IEnumerable<WorkOrder> WorkOrders
        {
            get { return workOrders; }
            set
            {
                workOrders = value;
                OnPropertyChanged(nameof(WorkOrders));
            }
        }
        public MaterialProfit MaterialProfitSelect
        {
            get { return materialProfitSelect; }
            set
            {
                materialProfitSelect = value;
                OnPropertyChanged(nameof(MaterialProfitSelect));
                if (MaterialProfitSelect == null)
                {
                    EnableMaterialProfit = false;
                }
                else
                {
                    EnableMaterialProfit = true;
                }
            }
        }
        public IEnumerable<MaterialProfit> MaterialProfits
        {
            get { return materialProfits; }
            set
            {
                materialProfits = value;
                OnPropertyChanged(nameof(MaterialProfits));
            }
        }
        public LabourProfit LabourProfitSelect
        {
            get { return labourProfitSelect; }
            set
            {
                labourProfitSelect = value;
                OnPropertyChanged(nameof(LabourProfitSelect));
                if (LabourProfitSelect == null)
                {
                    EnableLabourProfit = false;
                }
                else
                {
                    EnableLabourProfit = true;
                }
            }
        }
        public IEnumerable<LabourProfit> LabourProfits
        {
            get { return labourProfits; }
            set
            {
                labourProfits = value;
                OnPropertyChanged(nameof(LabourProfits));
            }
        }
        public Debts DebtSelect
        {
            get { return debtSelect; }
            set
            {
                debtSelect = value;
                OnPropertyChanged(nameof(DebtSelect));
                if (DebtSelect != null && DebtSelect.ReadOnly == false)
                {
                    EnableDebts = true;
                }
                else
                {
                    EnableDebts = false;
                }
            }
        }
        public IEnumerable<Debts> Debts
        {
            get { return debts; }
            set
            {
                debts = value;
                OnPropertyChanged(nameof(Debts));
            }
        }
        public Delivery DeliverySelect
        {
            get { return deliverySelect; }
            set
            {
                deliverySelect = value;
                OnPropertyChanged(nameof(DeliverySelect));
                if (DeliverySelect != null)
                {
                    DeliveryComboBoxSelect = DeliverySelect.NameComboBox;
                }
            }
        }
        public IEnumerable<Delivery> Deliveries
        {
            get { return deliveries; }
            set
            {
                deliveries = value;
                OnPropertyChanged(nameof(Deliveries));
                if (Deliveries != null)
                {
                    if (IsCheckedArchiveDelivery)
                    {
                        DeliveriesComboBox?.Clear();
                        DeliveriesComboBox = DeliveriesComboBoxArchiveGet();
                    }
                    else
                    {
                        DeliveriesComboBox?.Clear();
                        DeliveriesComboBox = DeliveriesComboBoxGet();
                    }
                }
            }
        }
        public List<string> DeliveriesComboBox
        {
            get { return deliveriesComboBox; }
            set
            {
                deliveriesComboBox = value;
                OnPropertyChanged(nameof(DeliveriesComboBox));
            }
        }
        public string DeliveryComboBoxSelect
        {
            get { return deliveryComboBoxSelect; }
            set
            {
                deliveryComboBoxSelect = value;
                OnPropertyChanged(nameof(DeliveryComboBoxSelect));
            }
        }
        public bool IsCheckedArchiveDelivery
        {
            get { return isCheckedArchiveDelivery; }
            set
            {
                isCheckedArchiveDelivery = value;
                OnPropertyChanged(nameof(IsCheckedArchiveDelivery));
                if (IsCheckedArchiveDelivery)
                {
                    DeliveriesComboBox?.Clear();
                    DeliveriesComboBox = DeliveriesComboBoxArchiveGet();
                    Deliveries = null;
                    LoadDeliveriesDB(CompanyName, true);
                }
                else
                {
                    DeliveriesComboBox?.Clear();
                    DeliveriesComboBox = DeliveriesComboBoxGet();
                    Deliveries = null;
                    LoadDeliveriesDB(CompanyName, false);
                }
            }
        }

        public bool EnableClient
        {
            get { return enableClient; }
            set
            {
                enableClient = value;
                OnPropertyChanged(nameof(EnableClient));
            }
        }
        public bool EnableQuota
        {
            get { return enableQuota; }
            set
            {
                enableQuota = value;
                OnPropertyChanged(nameof(EnableQuota));
            }
        }
        public bool EnableInvoice
        {
            get { return enableInvoice; }
            set
            {
                enableInvoice = value;
                OnPropertyChanged(nameof(EnableInvoice));
            }
        }
        public bool EnableWorkOrder
        {
            get { return enableWorkOrder; }
            set
            {
                enableWorkOrder = value;
                OnPropertyChanged(nameof(EnableWorkOrder));
            }
        }
        public bool EnableMaterialProfit
        {
            get { return enableMaterialProfit; }
            set
            {
                enableMaterialProfit = value;
                OnPropertyChanged(nameof(EnableMaterialProfit));
            }
        }
        public bool EnableLabourProfit
        {
            get { return enableLabourProfit; }
            set
            {
                enableLabourProfit = value;
                OnPropertyChanged(nameof(EnableLabourProfit));
            }
        }
        public bool EnableDebts
        {
            get { return enableDebts; }
            set
            {
                enableDebts = value;
                OnPropertyChanged(nameof(EnableDebts));
            }
        }

        public bool IsCheckedProfit
        {
            get { return isCheckedProfit; }
            set
            {
                isCheckedProfit = value;
                OnPropertyChanged(nameof(IsCheckedProfit));

                IsVisibleAmountReport = Visibility.Collapsed;
                IsVisibleDebtsReport = Visibility.Collapsed;
                IsVisibleExpensesReport = Visibility.Collapsed;
                IsVisiblePayrollReport = Visibility.Collapsed;
                IsVisibleProfitReport = Visibility.Collapsed;
                IsVisibleTotalReport = Visibility.Collapsed;
            }
        }
        public bool IsCheckedTotal
        {
            get { return isCheckedTotal; }
            set
            {
                isCheckedTotal = value;
                OnPropertyChanged(nameof(IsCheckedTotal));

                IsVisibleAmountReport = Visibility.Collapsed;
                IsVisibleDebtsReport = Visibility.Collapsed;
                IsVisibleExpensesReport = Visibility.Collapsed;
                IsVisiblePayrollReport = Visibility.Collapsed;
                IsVisibleProfitReport = Visibility.Collapsed;
                IsVisibleTotalReport = Visibility.Collapsed;
            }
        }
        public bool IsCheckedPayroll
        {
            get { return isCheckedPayroll; }
            set
            {
                isCheckedPayroll = value;
                OnPropertyChanged(nameof(IsCheckedPayroll));

                IsVisibleAmountReport = Visibility.Collapsed;
                IsVisibleDebtsReport = Visibility.Collapsed;
                IsVisibleExpensesReport = Visibility.Collapsed;
                IsVisiblePayrollReport = Visibility.Collapsed;
                IsVisibleProfitReport = Visibility.Collapsed;
                IsVisibleTotalReport = Visibility.Collapsed;
            }
        }
        public bool IsCheckedAmount
        {
            get { return isCheckedAmount; }
            set
            {
                isCheckedAmount = value;
                OnPropertyChanged(nameof(IsCheckedAmount));

                IsVisibleAmountReport = Visibility.Collapsed;
                IsVisibleDebtsReport = Visibility.Collapsed;
                IsVisibleExpensesReport = Visibility.Collapsed;
                IsVisiblePayrollReport = Visibility.Collapsed;
                IsVisibleProfitReport = Visibility.Collapsed;
                IsVisibleTotalReport = Visibility.Collapsed;
            }
        }
        public bool IsCheckedDebts
        {
            get { return isCheckedDebts; }
            set
            {
                isCheckedDebts = value;
                OnPropertyChanged(nameof(IsCheckedDebts));

                IsVisibleAmountReport = Visibility.Collapsed;
                IsVisibleDebtsReport = Visibility.Collapsed;
                IsVisibleExpensesReport = Visibility.Collapsed;
                IsVisiblePayrollReport = Visibility.Collapsed;
                IsVisibleProfitReport = Visibility.Collapsed;
                IsVisibleTotalReport = Visibility.Collapsed;
            }
        }
        public bool IsCheckedExpenses
        {
            get { return isCheckedExpenses; }
            set
            {
                isCheckedExpenses = value;
                OnPropertyChanged(nameof(IsCheckedExpenses));
                if (IsCheckedExpenses)
                {
                    IsVisibleDateSelect = Visibility.Collapsed;
                    IsVisibleDateSelectExpenses = Visibility.Visible;
                }
                else
                {
                    IsVisibleDateSelect = Visibility.Visible;
                    IsVisibleDateSelectExpenses = Visibility.Collapsed;
                }

                IsVisibleAmountReport = Visibility.Collapsed;
                IsVisibleDebtsReport = Visibility.Collapsed;
                IsVisibleExpensesReport = Visibility.Collapsed;
                IsVisiblePayrollReport = Visibility.Collapsed;
                IsVisibleProfitReport = Visibility.Collapsed;
                IsVisibleTotalReport = Visibility.Collapsed;
            }
        }
        public bool EnableReport
        {
            get { return enableReport; }
            set
            {
                enableReport = value;
                OnPropertyChanged(nameof(EnableReport));
            }
        }
        public Visibility IsVisibleMenuReport
        {
            get { return isVisibleMenuReport; }
            set
            {
                isVisibleMenuReport = value;
                OnPropertyChanged(nameof(IsVisibleMenuReport));
            }
        }
        public Visibility IsVisibleProfitReport
        {
            get { return isVisibleProfitReport; }
            set
            {
                isVisibleProfitReport = value;
                OnPropertyChanged(nameof(IsVisibleProfitReport));
            }
        }
        public Visibility IsVisibleTotalReport
        {
            get { return isVisibleTotalReport; }
            set
            {
                isVisibleTotalReport = value;
                OnPropertyChanged(nameof(IsVisibleTotalReport));
            }
        }
        public Visibility IsVisiblePayrollReport
        {
            get { return isVisiblePayrollReport; }
            set
            {
                isVisiblePayrollReport = value;
                OnPropertyChanged(nameof(IsVisiblePayrollReport));
            }
        }
        public Visibility IsVisibleAmountReport
        {
            get { return isVisibleAmountReport; }
            set
            {
                isVisibleAmountReport = value;
                OnPropertyChanged(nameof(IsVisibleAmountReport));
            }
        }
        public Visibility IsVisibleDebtsReport
        {
            get { return isVisibleDebtsReport; }
            set
            {
                isVisibleDebtsReport = value;
                OnPropertyChanged(nameof(IsVisibleDebtsReport));
            }
        }
        public Visibility IsVisibleExpensesReport
        {
            get { return isVisibleExpensesReport; }
            set
            {
                isVisibleExpensesReport = value;
                OnPropertyChanged(nameof(IsVisibleExpensesReport));
            }
        }
        public Visibility IsVisibleDateSelect
        {
            get { return isVisibleDateSelect; }
            set
            {
                isVisibleDateSelect = value;
                OnPropertyChanged(nameof(IsVisibleDateSelect));
            }
        }
        public Visibility IsVisibleDateSelectExpenses
        {
            get { return isVisibleDateSelectExpenses; }
            set
            {
                isVisibleDateSelectExpenses = value;
                OnPropertyChanged(nameof(IsVisibleDateSelectExpenses));
            }
        }
        public DateTime ReportDateFrom
        {
            get { return reportDateFrom; }
            set
            {
                reportDateFrom = value;
                OnPropertyChanged(nameof(ReportDateFrom));
                if (ReportDateFrom <= ReportDateTo)
                {
                    EnableReport = true;
                }
                else
                {
                    EnableReport = false;
                }
            }
        }
        public DateTime ReportDateTo
        {
            get { return reportDateTo; }
            set
            {
                reportDateTo = value;
                OnPropertyChanged(nameof(ReportDateTo));
                if (ReportDateFrom <= ReportDateTo)
                {
                    EnableReport = true;
                }
                else
                {
                    EnableReport = false;
                }
            }
        }
        public MaterialProfit MaterialReport
        {
            get { return materialReport; }
            set
            {
                materialReport = value;
                OnPropertyChanged(nameof(MaterialReport));
            }
        }
        public LabourProfit LabourReport
        {
            get { return labourReport; }
            set
            {
                labourReport = value;
                OnPropertyChanged(nameof(LabourReport));
            }
        }
        public List<Report> TotalReport
        {
            get { return totalReport; }
            set
            {
                totalReport = value;
                OnPropertyChanged(nameof(TotalReport));
            }
        }
        public List<ReportPayroll> ReportPayrolls
        {
            get { return reportPayrolls; }
            set
            {
                reportPayrolls = value;
                OnPropertyChanged(nameof(ReportPayrolls));
            }
        }
        public List<ReportPayrollToWork> ReportPayrollToWorks
        {
            get { return reportPayrollToWorks; }
            set
            {
                reportPayrollToWorks = value;
                OnPropertyChanged(nameof(ReportPayrollToWorks));
            }
        }
        public List<ReportAmount> ReportAmounts
        {
            get { return reportAmounts; }
            set
            {
                reportAmounts = value;
                OnPropertyChanged(nameof(ReportAmounts));
            }
        }
        public List<Debts> AmountDebtsReport
        {
            get { return amountDebtsReport; }
            set
            {
                amountDebtsReport = value;
                OnPropertyChanged(nameof(AmountDebtsReport));
            }
        }
        public List<Debts> PaymentDebtsReport
        {
            get { return paymentDebtsReport; }
            set
            {
                paymentDebtsReport = value;
                OnPropertyChanged(nameof(PaymentDebtsReport));
            }
        }
        public List<Expenses> ActiveExpenses
        {
            get { return activeExpenses; }
            set
            {
                activeExpenses = value;
                OnPropertyChanged(nameof(ActiveExpenses));
            }
        }
        public Expenses ActiveExpenseSelect
        {
            get { return activeExpenseSelect; }
            set
            {
                activeExpenseSelect = value;
                OnPropertyChanged(nameof(ActiveExpenseSelect));
            }
        }
        public List<Expenses> PaymentExpenses
        {
            get { return paymentExpenses; }
            set
            {
                paymentExpenses = value;
                OnPropertyChanged(nameof(PaymentExpenses));
            }
        }
        public Expenses PaymentExpenseSelect
        {
            get { return paymentExpenseSelect; }
            set
            {
                paymentExpenseSelect = value;
                OnPropertyChanged(nameof(PaymentExpenseSelect));
            }
        }
        public ReportTotal TotalReportForLabel
        {
            get { return totalReportForLabel; }
            set
            {
                totalReportForLabel = value;
                OnPropertyChanged(nameof(TotalReportForLabel));
            }
        }
        public string CountReport
        {
            get { return countReport; }
            set
            {
                countReport = value;
                OnPropertyChanged(nameof(CountReport));
            }
        }
        public decimal LabelAmountDebts
        {
            get { return labelAmountDebts; }
            set
            {
                labelAmountDebts = value;
                OnPropertyChanged(nameof(LabelAmountDebts));
            }
        }
        public decimal LabelPaymentDebts
        {
            get { return labelPaymentDebts; }
            set
            {
                labelPaymentDebts = value;
                OnPropertyChanged(nameof(LabelPaymentDebts));
            }
        }
        public decimal LabelActiveExpenses
        {
            get { return labelActiveExpenses; }
            set
            {
                labelActiveExpenses = value;
                OnPropertyChanged(nameof(LabelActiveExpenses));
            }
        }
        public decimal LabelPaymentExpenses
        {
            get { return labelPaymentExpenses; }
            set
            {
                labelPaymentExpenses = value;
                OnPropertyChanged(nameof(LabelPaymentExpenses));
            }
        }
        public int DateYearSelect
        {
            get { return dateYearSelect; }
            set
            {
                dateYearSelect = value;
                OnPropertyChanged(nameof(DateYearSelect));
            }
        }
        public List<int> DateYears
        {
            get { return dateYears; }
            set
            {
                dateYears = value;
                OnPropertyChanged(nameof(DateYears));
            }
        }
        public EnumMonths DateMonthSelect
        {
            get { return dateMonthSelect; }
            set
            {
                dateMonthSelect = value;
                OnPropertyChanged(nameof(DateMonthSelect));
            }
        }
        public List<EnumMonths> DateMonths
        {
            get { return dateMonths; }
            set
            {
                dateMonths = value;
                OnPropertyChanged(nameof(DateMonths));
            }
        }

        public bool IsCheckedCMO
        {
            get { return isCheckedCMO; }
            set
            {
                isCheckedCMO = value;
                OnPropertyChanged(nameof(IsCheckedCMO));
                if (IsCheckedCMO)
                {
                    CompanyName = "CMO";
                    StartLoadDB();
                }
            }
        }
        public bool IsCheckedLeveling
        {
            get { return isCheckedLeveling; }
            set
            {
                isCheckedLeveling = value;
                OnPropertyChanged(nameof(IsCheckedLeveling));
                if (IsCheckedLeveling)
                {
                    CompanyName = "Leveling";
                    StartLoadDB();
                }
            }
        }
        public Visibility IsVisibleCMO
        {
            get { return isVisibleCMO; }
            set
            {
                isVisibleCMO = value;
                OnPropertyChanged(nameof(IsVisibleCMO));
            }
        }
        public Visibility IsVisibleLeveling
        {
            get { return isVisibleLeveling; }
            set
            {
                isVisibleLeveling = value;
                OnPropertyChanged(nameof(IsVisibleLeveling));
            }
        }
        public string CompanyName
        {
            get { return companyName; }
            set
            {
                companyName = value;
                OnPropertyChanged(nameof(CompanyName));
            }
        }

        #endregion
        #region Private Command
        private Command _exitApp;

        private Command _loadFoto;
        private Command _viewFoto;
        //*****************************
        private Command _addClient;
        private Command _insClient;
        private Command _delClient;
        private Command _loadClient;
        private Command _searchClient;
        private Command _templateClient;
        //*********************************
        private Command _addQuotation;
        private Command _insQuotation;
        private Command _delQuotation;
        private Command _copyQuotation;
        private Command _printQuotation;
        private Command _paymentQuotation;
        private Command _loadQuotation;
        private Command _searchQuotation;
        private Command _templateQuotation;
        //*********************************
        private Command _printDelivery;
        private Command _printDriverDelivery;
        private Command _searchDelivery;
        private Command _dubleClickDelivery;
        //*********************************
        private Command _addInvoice;
        private Command _insInvoice;
        private Command _delInvoice;
        private Command _printInvoice;
        private Command _searchInvoice;
        //*********************************
        private Command _addWorkOrder;
        private Command _insWorkOrder;
        private Command _delWorkOrder;
        private Command _printWorkOrder;
        private Command _searchWorkOrder;
        //*********************************
        private Command _insMaterialProfit;
        private Command _printMaterialProfit;
        private Command _searchMaterialProfit;
        private Command _clickMaterialProfit;
        private Command _loadFotoMaterialProfit;
        private Command _viewFotoMaterialProfit;
        //*********************************
        private Command _insLabourProfit;
        private Command _printLabourProfit;
        private Command _searchLabourProfit;
        private Command _clickLabourProfit;
        //*********************************
        private Command _addDebts;
        private Command _insDebts;
        private Command _delDebts;
        private Command _paymentDebts;
        private Command _searchDebts;
        //**********************************
        private Command _loadReport;
        private Command _loadExpensesFromXls;
        private Command _loadExpenses;
        private Command _paymentExpensesActive;
        private Command _paymentExpensesPayment;
        private Command _templateExpenses;
        private Command _exportReport;
        //*************************************
        private Command _dicTypeOfClient;
        private Command _dicHearAboutsUs;
        private Command _dicQuotation;
        private Command _dicContractors;
        private Command _dicSupplier;
        private Command _dicDepth;
        //*************************************
        #endregion
        #region Public Command
        public Command ExitApp => _exitApp ?? (_exitApp = new Command(obj =>
        {
            ExitApplication();
        }));

        public Command LoadFoto => _loadFoto ?? (_loadFoto = new Command(obj =>
        {
            if (ClientSelect != null)
            {
                string[] files = OpenFileArray("Image files (*.png;*.jpeg;*.jpg;*.JPG)|*.png;*.jpeg;*.jpg;*.JPG");

                if (worker != null && worker.IsBusy) { return; }
                worker = new BackgroundWorker();
                worker.DoWork += worker_DoWork;
                worker.RunWorkerAsync();
                void worker_DoWork(object sender, DoWorkEventArgs e)
                {
                    ProgressStart();

                    // Основні затратні задачі
                    string SecondDir = Directory.GetCurrentDirectory() + "\\Foto";

                    if (files != null)
                    {
                        foreach (string fil in files)
                        {
                            File.Copy(fil, Path.Combine(SecondDir, ClientSelect.Id + "_client_" + Path.GetRandomFileName() + Path.GetExtension(fil)), true);
                        }
                    }
                    //************************

                    ProgressStop();
                }

            }

        }));
        public Command ViewFoto => _viewFoto ?? (_viewFoto = new Command(async obj =>
        {
            var displayRootRegistry = (Application.Current as App).displayRootRegistry;
            var foto = new FotoViewModel(ClientSelect.Id, "_client_*");
            await displayRootRegistry.ShowModalPresentation(foto);
        }));

        public Command AddClient => _addClient ?? (_addClient = new Command(async obj =>
        {
            var displayRootRegistry = (Application.Current as App).displayRootRegistry;
            var clientViewModel = new ClientViewModel(ref db, EnumClient.Add);
            clientViewModel.DateRegistration = DateTime.Today;
            await displayRootRegistry.ShowModalPresentation(clientViewModel);
            LoadClientsDB();
        }));
        public Command InsClient => _insClient ?? (_insClient = new Command(async obj =>
        {
            var displayRootRegistry = (Application.Current as App).displayRootRegistry;
            var clientViewModel = new ClientViewModel(ref db, EnumClient.Ins);
            if (ClientSelect != null)
            {
                var typeOfClient = db.DIC_TypeOfClients.FirstOrDefault(a => a.Name == ClientSelect.TypeOfClient);
                var hearAboutAs = db.DIC_HearAboutsUse.FirstOrDefault(a => a.Name == ClientSelect.HearAboutUs);
                clientViewModel.Id = ClientSelect.Id;
                clientViewModel.DateRegistration = ClientSelect.DateRegistration;
                clientViewModel.TypeOfClientSelect = typeOfClient;
                clientViewModel.CompanyName = ClientSelect.CompanyName;
                clientViewModel.PrimaryFirstName = ClientSelect.PrimaryFirstName;
                clientViewModel.PrimaryLastName = ClientSelect.PrimaryLastName;
                clientViewModel.PrimaryPhoneNumber = ClientSelect.PrimaryPhoneNumber;
                clientViewModel.PrimaryEmail = ClientSelect.PrimaryEmail;
                clientViewModel.SecondaryFirstName = ClientSelect.SecondaryFirstName;
                clientViewModel.SecondaryLastName = ClientSelect.SecondaryLastName;
                clientViewModel.SecondaryPhoneNumber = ClientSelect.SecondaryPhoneNumber;
                clientViewModel.SecondaryEmail = ClientSelect.SecondaryEmail;
                clientViewModel.AddressBillStreet = ClientSelect.AddressBillStreet;
                clientViewModel.AddressBillCity = ClientSelect.AddressBillCity;
                clientViewModel.AddressBillProvince = ClientSelect.AddressBillProvince;
                clientViewModel.AddressBillPostalCode = ClientSelect.AddressBillPostalCode;
                clientViewModel.AddressBillCountry = ClientSelect.AddressBillCountry;
                clientViewModel.AddressSiteStreet = ClientSelect.AddressSiteStreet;
                clientViewModel.AddressSiteCity = ClientSelect.AddressSiteCity;
                clientViewModel.AddressSiteProvince = ClientSelect.AddressSiteProvince;
                clientViewModel.AddressSitePostalCode = ClientSelect.AddressSitePostalCode;
                clientViewModel.AddressSiteCountry = ClientSelect.AddressSiteCountry;
                clientViewModel.HearAboutsUsSelect = hearAboutAs;
                clientViewModel.Specify = ClientSelect.Specify;
                clientViewModel.Notes = ClientSelect.Notes;

                await displayRootRegistry.ShowModalPresentation(clientViewModel);
                Clients = null;
                LoadClientsDB();
            }

        }));
        public Command DelClient => _delClient ?? (_delClient = new Command(obj =>
        {
            if (ClientSelect != null)
            {
                var quota = Quotations.Where(q => q.ClientId == ClientSelect.Id);

                foreach (var item in quota)
                {
                    var invoice = Invoices.Where(i => i.QuotaId == item.Id);
                    var material = db.MaterialQuotations.Where(m => m.QuotationId == item.Id);
                    var order = WorkOrders.Where(o => o.QuotaId == item.Id);
                    var receipt = db.Reciepts.Where(r => r.QuotaId == item.Id);
                    var payment = db.Payments.Where(p => p.QuotationId == item.Id);
                    var delivery = db.Deliveries.Where(d => d.QuotaId == item.Id);

                    foreach (var itemIn in invoice)         // Цю херню треба буде викинути і замутити обмеження при створенні Інвойса (один інвойс на одну квоту !!!)
                    {
                        var debts = Debts.Where(d => d.InvoiceId == itemIn.Id);
                        var matprof = MaterialProfits.FirstOrDefault(m => m.InvoiceId == itemIn.Id);
                        var mat = db.Materials.Where(m => m.MaterialProfitId == matprof.Id);

                        var labprof = LabourProfits.FirstOrDefault(l => l.InvoiceId == itemIn.Id);
                        var lab = db.Labours.Where(la => la.LabourProfitId == labprof.Id);
                        var labcontr = db.LabourContractors.Where(la => la.LabourProfitId == labprof.Id);

                        db.Debts.RemoveRange(debts);

                        db.Materials.RemoveRange(mat);
                        db.MaterialProfits.Remove(matprof);

                        db.Labours.RemoveRange(lab);
                        db.LabourContractors.RemoveRange(labcontr);
                        db.LabourProfits.Remove(labprof);
                    }

                    foreach (var ord in order)
                    {
                        var work = db.WorkOrder_Works.Where(w => w.WorkOrderId == ord.Id);
                        var accessories = db.WorkOrder_Accessories.Where(a => a.WorkOrderId == ord.Id);
                        var inst = db.WorkOrder_Installations.Where(i => i.WorkOrderId == ord.Id);
                        var con = db.WorkOrder_Contractors.Where(c => c.WorkOrderId == ord.Id);

                        db.WorkOrder_Contractors.RemoveRange(con);
                        db.WorkOrder_Installations.RemoveRange(inst);
                        db.WorkOrder_Accessories.RemoveRange(accessories);
                        db.WorkOrder_Works.RemoveRange(work);
                    }

                    foreach (var del in delivery)
                    {
                        var mat = db.DeliveryMaterials.Where(m => m.DeliveryId == del.Id);
                        db.DeliveryMaterials.RemoveRange(mat);
                    }

                    db.Payments.RemoveRange(payment);
                    db.Reciepts.RemoveRange(receipt);
                    db.WorkOrders.RemoveRange(order);
                    db.MaterialQuotations.RemoveRange(material);
                    db.Invoices.RemoveRange(invoice);
                    db.Deliveries.RemoveRange(delivery);

                    db.SaveChanges();
                }

                db.Quotations.RemoveRange(quota);
                db.SaveChanges();
                LoadQuotationsDB(CompanyName);

                db.Clients.Remove(ClientSelect);
                db.SaveChanges();
                LoadClientsDB();

                Deliveries = null;
                LoadDeliveriesDB(CompanyName, false);
                IsCheckedArchiveDelivery = false;
            }
        }));
        public Command LoadClient => _loadClient ?? (_loadClient = new Command(async obj =>
        {
            string path = OpenFile("Файл Excel|*.XLSX;*.XLS;*.XLSM");   // Вибираємо наш файл (метод OpenFile() описаний нижче)

            if (path == null) // Перевіряємо шлях до файлу на null
            {
                ProgressStop();
                return;
            }
            ProgressStart();
            db.Clients.Add(await GetClient(path));
            db.SaveChanges();
            ProgressStop();


            async Task<Client> GetClient(string pat)
            {

                Excel.Application ExcelApp = new Excel.Application();     // Створюємо додаток Excel
                Excel.Workbook ExcelWorkBook;                             // Створюємо книгу Excel
                Excel.Worksheet ExcelWorkSheet;                           // Створюємо лист Excel    
                Client client;

                try
                {
                    ExcelWorkBook = ExcelApp.Workbooks.Open(pat);                  // Відкриваємо файл Excel                
                    ExcelWorkSheet = ExcelWorkBook.ActiveSheet;                     // Відкриваємо активний Лист Excel                


                    client = await Task.Run(() => new Client()
                    {
                        DateRegistration = (DateTime.TryParse(ExcelApp.Cells[2, 3].Value?.ToString(), out DateTime res)) ? (res) : (DateTime.Today), // Для поля з датою треба просто "Value"
                        TypeOfClient = ExcelApp.Cells[3, 3].Value2?.ToString(),
                        CompanyName = ExcelApp.Cells[4, 3].Value2?.ToString(),

                        PrimaryFirstName = ExcelApp.Cells[5, 3].Value2?.ToString(),
                        PrimaryLastName = ExcelApp.Cells[6, 3].Value2?.ToString(),
                        PrimaryPhoneNumber = ExcelApp.Cells[7, 3].Value2?.ToString(),
                        PrimaryEmail = ExcelApp.Cells[8, 3].Value2?.ToString(),

                        SecondaryFirstName = ExcelApp.Cells[9, 3].Value2?.ToString(),
                        SecondaryLastName = ExcelApp.Cells[10, 3].Value2?.ToString(),
                        SecondaryPhoneNumber = ExcelApp.Cells[11, 3].Value2?.ToString(),
                        SecondaryEmail = ExcelApp.Cells[12, 3].Value2?.ToString(),

                        AddressBillStreet = ExcelApp.Cells[13, 3].Value2?.ToString(),
                        AddressBillCity = ExcelApp.Cells[14, 3].Value2?.ToString(),
                        AddressBillProvince = ExcelApp.Cells[15, 3].Value2?.ToString(),
                        AddressBillPostalCode = ExcelApp.Cells[16, 3].Value2?.ToString(),
                        AddressBillCountry = ExcelApp.Cells[17, 3].Value2?.ToString(),

                        AddressSiteStreet = ExcelApp.Cells[18, 3].Value2?.ToString(),
                        AddressSiteCity = ExcelApp.Cells[19, 3].Value2?.ToString(),
                        AddressSiteProvince = ExcelApp.Cells[20, 3].Value2?.ToString(),
                        AddressSitePostalCode = ExcelApp.Cells[21, 3].Value2?.ToString(),
                        AddressSiteCountry = ExcelApp.Cells[22, 3].Value2?.ToString(),

                        HearAboutUs = ExcelApp.Cells[23, 3].Value2?.ToString(),
                        Specify = ExcelApp.Cells[24, 3].Value2?.ToString(),
                        Notes = ExcelApp.Cells[25, 3].Value2?.ToString()
                    });


                    ExcelApp.Sheets[2].Cells[1, 20] = 1;

                    //ExcelApp.Visible = true;            // Робим Excel видимим, щоб користувач міг його закрити
                    //ExcelApp.UserControl = true;        // Передаємо керування користувачу


                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                    ExcelApp.Visible = true;
                    ExcelApp.UserControl = true;
                    client = null;
                }
                return client;
            }

        }));
        public Command SearchClient => _searchClient ?? (_searchClient = new Command(obj =>
        {
            string search = obj.ToString();
            if (search == "")
            {
                LoadClientsDB();
            }
            else
            {
                Clients = Clients.Where(n => n.FullSearch.ToUpper().Contains(search.ToUpper()));
            }
        }));
        public Command TemplateClient => _templateClient ?? (_templateClient = new Command(obj =>
        {
            if (worker != null && worker.IsBusy) { return; }

            worker = new BackgroundWorker();
            worker.DoWork += worker_DoWork;
            worker.RunWorkerAsync();
            void worker_DoWork(object sender, DoWorkEventArgs e)
            {
                ProgressStart();
                try
                {
                    // Основні затратні задачі
                    Excel.Application ExcelApp = new Excel.Application();
                    Excel.Workbook ExcelWorkBook;
                    ExcelWorkBook = ExcelApp.Workbooks.Open(Environment.CurrentDirectory + "\\Blanks\\TemplateClient");   //Вказуємо шлях до шаблону "\\Blanks\\TemplateClient.xltm"

                    var typeClient = db.DIC_TypeOfClients.Where(c => c.Id > 0);
                    var howClient = db.DIC_HearAboutsUse.Where(h => h.Id > 0);
                    int i = 2;
                    foreach (var item in typeClient)
                    {
                        ExcelApp.Sheets[2].Cells[i, 1] = item.Name;
                        i++;
                    }

                    i = 2;
                    foreach (var item in howClient)
                    {
                        ExcelApp.Sheets[2].Cells[i, 2] = item.Name;
                        i++;
                    }

                    ExcelApp.Visible = true;           // Робим книгу видимою
                    ExcelApp.UserControl = true;       // Передаємо керування користувачу  
                                                        //************************
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error message: " + Environment.NewLine +
                                           ex.Message + Environment.NewLine + Environment.NewLine +
                                           "StackTrace message: " + Environment.NewLine +
                                           ex.StackTrace, "Warning !!!");                    
                }
                ProgressStop();
            }

        }));

        //**********************************
        public Command AddQuotation => _addQuotation ?? (_addQuotation = new Command(async obj =>
        {
            var displayRootRegistry = (Application.Current as App).displayRootRegistry;
            var quotationViewModel = new QuotationViewModel(ref db, EnumClient.Add, QuotationSelect, CompanyName);
            quotationViewModel.Clients = Clients;
            await displayRootRegistry.ShowModalPresentation(quotationViewModel);
            QuotationSelect = quotationViewModel.QuotaSelect;

            Calculate(quotationViewModel.QuotaId);

        }));
        public Command InsQuotation => _insQuotation ?? (_insQuotation = new Command(async obj =>
        {
            if (QuotationSelect != null)
            {
                decimal start = QuotationSelect.InvoiceGrandTotal; // Початкові дані для  порівняння
                var quotaStart = QuotationSelect;

                var displayRootRegistry = (Application.Current as App).displayRootRegistry;
                var quotationViewModel = new QuotationViewModel(ref db, EnumClient.Ins, QuotationSelect, CompanyName);
                quotationViewModel.Clients = Clients;
                quotationViewModel.ClientSelect = db.Clients.FirstOrDefault(c => c.Id == QuotationSelect.ClientId);
                await displayRootRegistry.ShowModalPresentation(quotationViewModel);
                Calculate(QuotationSelect.Id);

                QuotationSelect = Quotations.FirstOrDefault(q => q.Id == quotaStart.Id);
                var quotaFinish = QuotationSelect;  // Кінцеві дані для порівняння (що було і що стало)
                decimal finish = QuotationSelect.InvoiceGrandTotal;
                if (start != finish)
                {
                    DelDeliveri(quotaFinish);
                    var paid = db.Payments.Where(p => p.QuotationId == quotaFinish.Id);

                    if (paid != null)
                    {
                        CalculatePayment(QuotationSelect);         // Перераховуєм Баланс                        
                    }
                }
            }
        }));
        public Command DelQuotation => _delQuotation ?? (_delQuotation = new Command(obj =>
        {
            if (QuotationSelect != null)
            {
                var material = db.MaterialQuotations.Where(m => m.QuotationId == QuotationSelect.Id);
                var invoice = Invoices.Where(i => i.QuotaId == QuotationSelect.Id);
                var order = WorkOrders.Where(o => o.QuotaId == QuotationSelect.Id);
                var receipt = db.Reciepts.Where(r => r.QuotaId == QuotationSelect.Id);
                var pay = db.Payments.Where(p => p.QuotationId == QuotationSelect.Id);
                var delivery = db.Deliveries.Where(d => d.QuotaId == QuotationSelect.Id);

                foreach (var item in delivery)
                {
                    var mat = db.DeliveryMaterials.Where(m => m.DeliveryId == item.Id);
                    db.DeliveryMaterials.RemoveRange(mat);
                }

                foreach (var item in invoice)         // Цю херню треба буде викинути і замутити обмеження при створенні Інвойса (один інвойс на одну квоту !!!)
                {
                    var matprof = MaterialProfits.FirstOrDefault(m => m.InvoiceId == item.Id);
                    var mat = db.Materials.Where(m => m.MaterialProfitId == matprof.Id);

                    var labprof = LabourProfits.FirstOrDefault(l => l.InvoiceId == item.Id);
                    var lab = db.Labours.Where(la => la.LabourProfitId == labprof.Id);
                    var labcontr = db.LabourContractors.Where(la => la.LabourProfitId == labprof.Id);

                    var debts = Debts.Where(d => d.InvoiceId == item.Id);

                    db.Materials.RemoveRange(mat);
                    db.MaterialProfits.Remove(matprof);

                    db.Labours.RemoveRange(lab);
                    db.LabourContractors.RemoveRange(labcontr);
                    db.LabourProfits.Remove(labprof);

                    db.Debts.RemoveRange(debts);
                }

                foreach (var item in order)
                {
                    var work = db.WorkOrder_Works.Where(w => w.WorkOrderId == item.Id);
                    var accessories = db.WorkOrder_Accessories.Where(a => a.WorkOrderId == item.Id);
                    var inst = db.WorkOrder_Installations.Where(ins => ins.WorkOrderId == item.Id);
                    var con = db.WorkOrder_Contractors.Where(c => c.WorkOrderId == item.Id);
                    db.WorkOrder_Contractors.RemoveRange(con);
                    db.WorkOrder_Installations.RemoveRange(inst);
                    db.WorkOrder_Accessories.RemoveRange(accessories);
                    db.WorkOrder_Works.RemoveRange(work);
                }
                db.Deliveries.RemoveRange(delivery);
                db.WorkOrders.RemoveRange(order);
                db.Reciepts.RemoveRange(receipt);
                db.Payments.RemoveRange(pay);
                db.Invoices.RemoveRange(invoice);
                db.MaterialQuotations.RemoveRange(material);
                db.Quotations.Remove(QuotationSelect);
                db.SaveChanges();

                Quotations = null;
                LoadQuotationsDB(CompanyName);

                Deliveries = null;
                LoadDeliveriesDB(CompanyName, false);
                IsCheckedArchiveDelivery = false;

                WorkOrders = null;
                LoadWorkOrdersDB(CompanyName);

                Invoices = null;
                LoadInvoicesDB(CompanyName);

                MaterialProfits = null;
                LoadMaterialProfitsDB(CompanyName);

                LabourProfits = null;
                LoadLabourProfitsDB(CompanyName);

                Debts = null;
                LoadDebtsDB(CompanyName);
            }
        }));
        public Command CopyQuotation => _copyQuotation ?? (_copyQuotation = new Command(async obj =>
        {
            if (QuotationSelect != null)
            {
                var displayRootRegistry = (Application.Current as App).displayRootRegistry;

                var select = new SelectClientViewModel();
                select.LastId = QuotationSelect.ClientId;
                select.Clients = Clients;
                select.ClientSelect = Clients.FirstOrDefault(c => c.Id == QuotationSelect.ClientId); //db.Clients.FirstOrDefault(c => c.Id == QuotationSelect.ClientId);

                await displayRootRegistry.ShowModalPresentation(select);

                if (select.PressButton)
                {
                    Quotation quota = new Quotation()
                    {
                        ClientId = select.ClientSelect.Id,
                        NumberClient = select.ClientSelect.NumberClient,
                        FirstName = select.ClientSelect.PrimaryFirstName,
                        LastName = select.ClientSelect.PrimaryLastName,
                        PhoneNumber = select.ClientSelect.PrimaryPhoneNumber,
                        Email = select.ClientSelect.PrimaryEmail,
                        QuotaDate = DateTime.Today,
                        PrefixNumberQuota = QuotationSelect.PrefixNumberQuota,
                        AmountPaidByCreditCard = QuotationSelect.AmountPaidByCreditCard,

                        FinancingYesNo = QuotationSelect.FinancingYesNo,
                        JobDescription = QuotationSelect.JobDescription,
                        JobNote = QuotationSelect.JobNote,
                        LabourDiscountAmount = QuotationSelect.LabourDiscountAmount,
                        LabourDiscountYN = QuotationSelect.LabourDiscountYN,
                        LabourSubtotal = QuotationSelect.LabourSubtotal,
                        MaterialDiscountAmount = QuotationSelect.MaterialDiscountAmount,
                        MaterialDiscountYN = QuotationSelect.MaterialDiscountYN,
                        MaterialSubtotal = QuotationSelect.MaterialSubtotal,

                    };

                    db.Quotations.Add(quota);
                    db.SaveChanges();


                    quota = Quotations.FirstOrDefault(q => q.Id == QuotationSelect.Id);  // попередня квота
                    var mat = db.MaterialQuotations.Where(m => m.QuotationId == quota.Id);
                    var pay = db.Payments.Where(p => p.QuotationId == quota.Id);
                    QuotationSelect = Quotations.OrderByDescending(q => q.Id).FirstOrDefault(); // новостворена квота

                    foreach (var item in mat)
                    {
                        MaterialQuotation material = new MaterialQuotation()
                        {
                            Description = item.Description,
                            Groupe = item.Groupe,
                            Item = item.Item,
                            Price = item.Price,
                            Quantity = item.Quantity,
                            QuotationId = QuotationSelect.Id,
                            Rate = item.Rate
                        };
                        db.MaterialQuotations.Add(material);
                        db.SaveChanges();
                    }

                    foreach (var p in pay)
                    {
                        Payment payment = new Payment()
                        {
                            NumberPayment = p.NumberPayment,
                            PaymentDatePaid = p.PaymentDatePaid,
                            PaymentAmountPaid = p.PaymentAmountPaid,
                            PaymentMethod = p.PaymentMethod,
                            Balance = p.Balance,
                            QuotationId = QuotationSelect.Id
                        };
                        db.Payments.Add(payment);
                        db.SaveChanges();

                        Reciept rec = new Reciept()
                        {
                            NumberQuota = QuotationSelect.NumberQuota,
                            PayNumber = payment.Id,
                            QuotaId = QuotationSelect.Id
                        };
                        db.Reciepts.Add(rec);
                        db.SaveChanges();
                    }

                    Calculate(QuotationSelect.Id);
                }
            }
        }));
        public Command PrintQuotation => _printQuotation ?? (_printQuotation = new Command(obj =>
        {
            try
            {
                if (CompanyName == "CMO")
                {
                    QuotaPrintToExcelCMO("\\Blanks\\QuotaPDF");   //"\\Blanks\\QuotaPDF.xltm"
                }
                else
                {
                    QuotaPrintToExcelNL("\\Blanks\\NLQuotaPDF");   // "\\Blanks\\NLQuotaPDF.xltm"
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error message: " + Environment.NewLine +
                                       ex.Message + Environment.NewLine + Environment.NewLine +
                                       "StackTrace message: " + Environment.NewLine +
                                       ex.StackTrace, "Warning !!!");
            }
        }));
        public Command PaymentQuotation => _paymentQuotation ?? (_paymentQuotation = new Command(async obj =>
        {
            if (QuotationSelect != null)
            {

                int? quotaId = QuotationSelect.Id;
                var displayRootRegistry = (Application.Current as App).displayRootRegistry;
                var payment = new PaymentViewModel(ref db, QuotationSelect);
                await displayRootRegistry.ShowModalPresentation(payment);
                Quotations = null;
                LoadQuotationsDB(CompanyName);

                var quota = Quotations.Where(q => q.Id == quotaId);
                if (quota != null)
                {
                    var quotaItem = quota.FirstOrDefault(i => i.Id == quotaId);
                    var order = WorkOrders.FirstOrDefault(w => w.QuotaId == quotaId);
                    var invoice = Invoices.FirstOrDefault(i => i.QuotaId == quotaItem.Id);
                    var delivery = Deliveries.FirstOrDefault(d => d.QuotaId == quotaItem.Id);

                    // Creating Delivery
                    if (quotaItem.ActivQuota && delivery == null)
                    {
                        displayRootRegistry = (Application.Current as App).displayRootRegistry;
                        var message = new MessageViewModel(200, 410, "   Quote status changed to Active !!!" + Environment.NewLine + "Do you want to create a 'Delivery' ? ");
                        await displayRootRegistry.ShowModalPresentation(message);
                        if (message.PressOk)
                        {
                            AddDelivery(quotaItem);
                        }
                    }

                    // Creating WorkOrder
                    if (quotaItem.ActivQuota && order == null)
                    {
                        displayRootRegistry = (Application.Current as App).displayRootRegistry;
                        var message = new MessageViewModel(200, 410, "   Quote status changed to Active !!!" + Environment.NewLine + "Do you want to create a 'Work Order' ? ");
                        await displayRootRegistry.ShowModalPresentation(message);
                        if (message.PressOk)
                        {
                            var work = new WorkOrderViewModel(ref db, EnumClient.Add, null);
                            work.Quotations = quota;
                            work.QuotationSelect = quota.FirstOrDefault(q => q.Id == quotaId);
                            work.CreatOrder.Execute("");
                        }
                        LoadWorkOrdersDB(CompanyName);
                    }
                    // Creating Invoice                    
                    if (quotaItem.PaidQuota && invoice == null && order != null)
                    {
                        displayRootRegistry = (Application.Current as App).displayRootRegistry;
                        var message = new MessageViewModel(200, 410, "   The Quote is fully paid !!!" + Environment.NewLine + "Do you want to move her to the 'Completed Jobs' ?");
                        await displayRootRegistry.ShowModalPresentation(message);
                        if (message.PressOk)
                        {
                            var temp = Clients.FirstOrDefault(q => q.Id == quotaItem.ClientId);
                            Invoice inv = new Invoice()
                            {
                                DateInvoice = DateTime.Today,
                                NumberQuota = quotaItem.NumberQuota,
                                OrderNumber = "",
                                UpNumber = "",
                                QuotaId = quotaItem.Id,
                                FirstName = temp.PrimaryFirstName,
                                LastName = temp.PrimaryLastName,
                                PhoneNumber = temp.PrimaryPhoneNumber,
                                Email = temp.PrimaryEmail,
                                CompanyName = quotaItem.CompanyName
                            };

                            var deliveryToArchive = db.Deliveries.Where(d => d.QuotaId == quotaId);
                            foreach (var item in deliveryToArchive)
                            {
                                item.IsArchive = true;
                                db.Entry(item).State = EntityState.Modified;
                            }
                            db.Invoices.Add(inv);
                            db.SaveChanges();

                            Deliveries = null;
                            LoadDeliveriesDB(CompanyName, false);
                            IsCheckedArchiveDelivery = false;

                            int? NewInvoiceId = Invoices.OrderByDescending(q => q.Id).FirstOrDefault().Id;
                            AddMaterialProfit(NewInvoiceId);
                            AddLabourProfit(NewInvoiceId);
                            AddDebtsTab(NewInvoiceId);
                            var material = MaterialProfits.FirstOrDefault(m => m.InvoiceId == NewInvoiceId);
                            var labour = LabourProfits.FirstOrDefault(l => l.InvoiceId == NewInvoiceId);
                            MaterialProfitCalculate(material);
                            LabourProfitCalculate(labour);
                            InvoiceCalculate(NewInvoiceId);
                        }
                    }
                }
            }
        }));
        public Command LoadQuotation => _loadQuotation ?? (_loadQuotation = new Command(obj =>
        {
            if (ClientSelect != null)
            {
                string path = OpenFile("Файл Excel|*.XLSX;*.XLS;*.XLSM");   // Вибираємо наш файл (метод OpenFile() описаний нижче)

                if (path == null) // Перевіряємо шлях до файлу на null
                {
                    return;
                }


                if (worker != null && worker.IsBusy) { return; }
                worker = new BackgroundWorker();
                worker.DoWork += worker_DoWork;
                worker.RunWorkerAsync();
                void worker_DoWork(object sender, DoWorkEventArgs e)
                {
                    ProgressStart();

                    // Основні затратні задачі

                    Excel.Application ExcelApp = new Excel.Application();     // Створюємо додаток Excel
                    Excel.Workbook ExcelWorkBook;                             // Створюємо книгу Excel
                    Excel.Worksheet ExcelWorkSheet;                           // Створюємо лист Excel

                    try
                    {

                        ExcelWorkBook = ExcelApp.Workbooks.Open(path);                  // Відкриваємо файл Excel                
                        ExcelWorkSheet = ExcelWorkBook.ActiveSheet;                     // Відкриваємо активний Лист Excel

                        Client client = ClientSelect;

                        Quotation quota = new Quotation()
                        {
                            QuotaDate = DateTime.Today,
                            PrefixNumberQuota = "Q",
                            ClientId = client.Id,
                            NumberClient = client.NumberClient,
                            FirstName = client.PrimaryFirstName,
                            LastName = client.PrimaryLastName,
                            PhoneNumber = client.PrimaryPhoneNumber,
                            Email = client.PrimaryEmail,
                            JobDescription = ExcelApp.Cells[2, 2].Value2?.ToString(),
                            JobNote = ExcelApp.Cells[3, 2].Value2?.ToString(),
                            FinancingYesNo = false,

                        };
                        db.Quotations.Add(quota);
                        db.SaveChanges();
                        var quotaSelect = db.Quotations.Local.ToBindingList().OrderByDescending(q => q.Id).FirstOrDefault();
                        QuotationSelect = quotaSelect;


                        // "FLOORING"
                        for (int i = 7; i <= 14; i++)
                        {
                            string nameG = "FLOORING";
                            string nameI = ExcelApp.Cells[i, 1].Value2?.ToString();
                            string nameD = ExcelApp.Cells[i, 2].Value2?.ToString();
                            string nameQ = ExcelApp.Cells[i, 3].Value2?.ToString();

                            if (nameQ != null)
                            {
                                var groupe = db.DIC_GroupeQuotations.FirstOrDefault(g => g.NameGroupe == nameG);
                                var item = db.DIC_ItemQuotations.FirstOrDefault(itm => itm.GroupeId == groupe.Id && itm.Name == nameI);
                                var des = db.DIC_DescriptionQuotations.FirstOrDefault(d => d.ItemId == item.Id && d.Name == nameD);

                                decimal rate = (des != null) ? (des.Price) : (0m);
                                decimal quantity = (decimal.TryParse(nameQ, out decimal res)) ? (res) : (0m);

                                MaterialQuotation material = new MaterialQuotation()
                                {
                                    QuotationId = quotaSelect.Id,
                                    Groupe = nameG,
                                    Item = nameI,
                                    Description = nameD,
                                    Quantity = quantity,
                                    Rate = rate,
                                    Price = decimal.Round(rate * quantity, 2)
                                };
                                db.MaterialQuotations.Add(material);
                                db.SaveChanges();
                            }
                        }
                        // "ACCESSORIES"
                        for (int i = 16; i <= 31; i++)
                        {
                            string nameG = "ACCESSORIES";
                            string nameI = ExcelApp.Cells[i, 1].Value2?.ToString();
                            string nameD = ExcelApp.Cells[i, 2].Value2?.ToString();
                            string nameQ = ExcelApp.Cells[i, 3].Value2?.ToString();

                            if (nameQ != null)
                            {
                                var groupe = db.DIC_GroupeQuotations.FirstOrDefault(g => g.NameGroupe == nameG);
                                var item = db.DIC_ItemQuotations.FirstOrDefault(itm => itm.GroupeId == groupe.Id && itm.Name == nameI);
                                var des = db.DIC_DescriptionQuotations.FirstOrDefault(d => d.ItemId == item.Id && d.Name == nameD);

                                decimal rate = (des != null) ? (des.Price) : (0m);
                                decimal quantity = (decimal.TryParse(nameQ, out decimal res)) ? (res) : (0m);

                                MaterialQuotation material = new MaterialQuotation()
                                {
                                    QuotationId = quotaSelect.Id,
                                    Groupe = nameG,
                                    Item = nameI,
                                    Description = nameD,
                                    Quantity = quantity,
                                    Rate = rate,
                                    Price = decimal.Round(rate * quantity, 2)
                                };
                                db.MaterialQuotations.Add(material);
                                db.SaveChanges();
                            }
                        }
                        // "INSTALLATION"
                        for (int i = 34; i <= 41; i++)
                        {
                            string nameG = "INSTALLATION";
                            string nameI = ExcelApp.Cells[i, 1].Value2?.ToString();
                            string nameD = ExcelApp.Cells[i, 2].Value2?.ToString();
                            string nameQ = ExcelApp.Cells[i, 3].Value2?.ToString();

                            if (nameQ != null)
                            {
                                var groupe = db.DIC_GroupeQuotations.FirstOrDefault(g => g.NameGroupe == nameG);
                                var item = db.DIC_ItemQuotations.FirstOrDefault(itm => itm.GroupeId == groupe.Id && itm.Name == nameI);
                                var des = db.DIC_DescriptionQuotations.FirstOrDefault(d => d.ItemId == item.Id && d.Name == nameD);

                                decimal rate = (des != null) ? (des.Price) : (0m);
                                decimal quantity = (decimal.TryParse(nameQ, out decimal res)) ? (res) : (0m);

                                MaterialQuotation material = new MaterialQuotation()
                                {
                                    QuotationId = quotaSelect.Id,
                                    Groupe = nameG,
                                    Item = nameI,
                                    Description = nameD,
                                    Quantity = quantity,
                                    Rate = rate,
                                    Price = decimal.Round(rate * quantity, 2)
                                };
                                db.MaterialQuotations.Add(material);
                                db.SaveChanges();
                            }
                        }
                        // "DEMOLITION"
                        for (int i = 43; i <= 47; i++)
                        {
                            string nameG = "DEMOLITION";
                            string nameI = ExcelApp.Cells[i, 1].Value2?.ToString();
                            string nameD = ExcelApp.Cells[i, 2].Value2?.ToString();
                            string nameQ = ExcelApp.Cells[i, 3].Value2?.ToString();

                            if (nameQ != null)
                            {
                                var groupe = db.DIC_GroupeQuotations.FirstOrDefault(g => g.NameGroupe == nameG);
                                var item = db.DIC_ItemQuotations.FirstOrDefault(itm => itm.GroupeId == groupe.Id && itm.Name == nameI);
                                var des = db.DIC_DescriptionQuotations.FirstOrDefault(d => d.ItemId == item.Id && d.Name == nameD);

                                decimal rate = (des != null) ? (des.Price) : (0m);
                                decimal quantity = (decimal.TryParse(nameQ, out decimal res)) ? (res) : (0m);

                                MaterialQuotation material = new MaterialQuotation()
                                {
                                    QuotationId = quotaSelect.Id,
                                    Groupe = nameG,
                                    Item = nameI,
                                    Description = nameD,
                                    Quantity = quantity,
                                    Rate = rate,
                                    Price = decimal.Round(rate * quantity, 2)
                                };
                                db.MaterialQuotations.Add(material);
                                db.SaveChanges();
                            }
                        }
                        // "OPTIONAL SERVICES"
                        for (int i = 49; i <= 55; i++)
                        {
                            string nameG = "OPTIONAL SERVICES";
                            string nameI = ExcelApp.Cells[i, 1].Value2?.ToString();
                            string nameD = ExcelApp.Cells[i, 2].Value2?.ToString();
                            string nameQ = ExcelApp.Cells[i, 3].Value2?.ToString();

                            if (nameQ != null)
                            {
                                var groupe = db.DIC_GroupeQuotations.FirstOrDefault(g => g.NameGroupe == nameG);
                                var item = db.DIC_ItemQuotations.FirstOrDefault(itm => itm.GroupeId == groupe.Id && itm.Name == nameI);
                                var des = db.DIC_DescriptionQuotations.FirstOrDefault(d => d.ItemId == item.Id && d.Name == nameD);

                                decimal rate = (des != null) ? (des.Price) : (0m);
                                decimal quantity = (decimal.TryParse(nameQ, out decimal res)) ? (res) : (0m);

                                MaterialQuotation material = new MaterialQuotation()
                                {
                                    QuotationId = quotaSelect.Id,
                                    Groupe = nameG,
                                    Item = nameI,
                                    Description = nameD,
                                    Quantity = quantity,
                                    Rate = rate,
                                    Price = decimal.Round(rate * quantity, 2)
                                };
                                db.MaterialQuotations.Add(material);
                                db.SaveChanges();
                            }
                        }
                        // "FLOORING DELIVERY"
                        for (int i = 57; i <= 60; i++)
                        {
                            string nameG = "FLOORING DELIVERY";
                            string nameI = ExcelApp.Cells[i, 1].Value2?.ToString();
                            string nameD = ExcelApp.Cells[i, 2].Value2?.ToString();
                            string nameQ = ExcelApp.Cells[i, 3].Value2?.ToString();

                            if (nameQ != null)
                            {
                                var groupe = db.DIC_GroupeQuotations.FirstOrDefault(g => g.NameGroupe == nameG);
                                var item = db.DIC_ItemQuotations.FirstOrDefault(itm => itm.GroupeId == groupe.Id && itm.Name == nameI);
                                var des = db.DIC_DescriptionQuotations.FirstOrDefault(d => d.ItemId == item.Id && d.Name == nameD);

                                decimal rate = (des != null) ? (des.Price) : (0m);
                                decimal quantity = (decimal.TryParse(nameQ, out decimal res)) ? (res) : (0m);

                                MaterialQuotation material = new MaterialQuotation()
                                {
                                    QuotationId = quotaSelect.Id,
                                    Groupe = nameG,
                                    Item = nameI,
                                    Description = nameD,
                                    Quantity = quantity,
                                    Rate = rate,
                                    Price = decimal.Round(rate * quantity, 2)
                                };
                                db.MaterialQuotations.Add(material);
                                db.SaveChanges();
                            }
                        }

                        ExcelApp.Sheets[2].Cells[1, 20] = 1;

                        //ExcelApp.Visible = true;            // Робим Excel видимим, щоб користувач міг його закрити
                        //ExcelApp.UserControl = true;        // Передаємо керування користувачу

                        Calculate(quotaSelect.Id);

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                        ExcelApp.Visible = true;
                        ExcelApp.UserControl = true;
                        Calculate(QuotationSelect.Id);
                    }
                    //************************

                    ProgressStop();
                }
            }
        }));
        public Command SearchQuotation => _searchQuotation ?? (_searchQuotation = new Command(obj =>
        {
            string search = obj.ToString();
            if (search == "")
            {
                LoadQuotationsDB(CompanyName);
            }
            else
            {
                Quotations = Quotations.Where(n => n.CompanyName == CompanyName && n.FullSearch.ToUpper().Contains(search.ToUpper())).OrderBy(q => q.SortingQuota).ThenBy(q => q.Id);
            }
        }));
        public Command TemplateQuotation => _templateQuotation ?? (_templateQuotation = new Command(obj =>
        {
            if (worker != null && worker.IsBusy) { return; }
            worker = new BackgroundWorker();
            worker.DoWork += worker_DoWork;
            worker.RunWorkerAsync();
            void worker_DoWork(object sender, DoWorkEventArgs e)
            {
                ProgressStart();
                try
                {
                    // Основні затратні задачі
                    Excel.Application ExcelApp = new Excel.Application();
                    Excel.Workbook ExcelWorkBook;
                    ExcelWorkBook = ExcelApp.Workbooks.Open(Environment.CurrentDirectory + "\\Blanks\\TemplateQuota");   //Вказуємо шлях до шаблону  "\\Blanks\\TemplateQuota.xltm"

                    var groups = db.DIC_GroupeQuotations.Where(g => g.Id > 0);
                    var dic_item = db.DIC_ItemQuotations.Where(g => g.Id > 0);
                    var dic_des = db.DIC_DescriptionQuotations.Where(g => g.Id > 0);

                    int i = 2;
                    int its = 2;
                    foreach (var group in groups)
                    {
                        var items = dic_item.Where(it => it.GroupeId == group.Id).OrderBy(it => it.Name);
                        foreach (var item in items)
                        {
                            ExcelApp.Sheets[2].Cells[its, 1] = item.Name;
                            ExcelApp.Sheets[2].Cells[its, 2] = group.NameGroupe;
                            its++;

                            var descroptions = dic_des.Where(d => d.ItemId == item.Id);
                            foreach (var description in descroptions)
                            {
                                ExcelApp.Sheets[2].Cells[i, 4] = group.NameGroupe;
                                ExcelApp.Sheets[2].Cells[i, 5] = item.Name;
                                ExcelApp.Sheets[2].Cells[i, 6] = description.Name;
                                i++;
                            }
                        }
                    }
                    groups = null;
                    dic_item = null;
                    dic_des = null;

                    ExcelApp.Visible = true;           // Робим книгу видимою
                    ExcelApp.UserControl = true;       // Передаємо керування користувачу 
                }                                      //************************
                catch (Exception ex)
                {
                    MessageBox.Show("Error message: " + Environment.NewLine +
                                           ex.Message + Environment.NewLine + Environment.NewLine +
                                           "StackTrace message: " + Environment.NewLine +
                                           ex.StackTrace, "Warning !!!");
                }

                ProgressStop();
            }
        }));

        //**********************************
        public Command PrintDelivery => _printDelivery ?? (_printDelivery = new Command(obj =>
        {
            try
            {
                if (CompanyName == "CMO")
                {
                    DeliveryPrintToExcelCMO("\\Blanks\\ListOfSuppliesPDF");   //  "\\Blanks\\ListOfSuppliesPDF.xltm"
                }
                else
                {
                    DeliveryPrintToExcelNL("\\Blanks\\NLListOfSuppliesPDF");   //  "\\Blanks\\NLListOfSuppliesPDF.xltm"
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error message: " + Environment.NewLine +
                                       ex.Message + Environment.NewLine + Environment.NewLine +
                                       "StackTrace message: " + Environment.NewLine +
                                       ex.StackTrace, "Warning !!!");
            }
        }));
        public Command PrintDriverDelivery => _printDriverDelivery ?? (_printDriverDelivery = new Command(obj =>
        {
            try
            {

                if (CompanyName == "CMO")
                {
                    DeliveryDriverPrintYoExcelCMO("\\Blanks\\DeliveryPDF");    // "\\Blanks\\DeliveryPDF.xltm"
                }
                else
                {
                    DeliveryDriverPrintYoExcelNL("\\Blanks\\NLDeliveryPDF");   //  "\\Blanks\\NLDeliveryPDF.xltm"
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error message: " + Environment.NewLine +
                                       ex.Message + Environment.NewLine + Environment.NewLine +
                                       "StackTrace message: " + Environment.NewLine +
                                       ex.StackTrace, "Warning !!!");
            }
        }));
        public Command SearchDelivery => _searchDelivery ?? (_searchDelivery = new Command(obj =>
        {
            string search = obj.ToString();
            if (search == "")
            {
                LoadDeliveriesDB(CompanyName, false);
            }
            else
            {
                Deliveries = Deliveries.Where(n => n.CompanyName == CompanyName && n.FullSearch.ToUpper().Contains(search.ToUpper()));
            }
        }));
        public Command DubleClickDelivery => _dubleClickDelivery ?? (_dubleClickDelivery = new Command(async obj =>
        {
            if (DeliverySelect != null)
            {
                var displayRootRegistry = (Application.Current as App).displayRootRegistry;
                var deliveryVM = new DeliveryViewModel();
                deliveryVM.OrderNumber = DeliverySelect?.OrderNumber;
                deliveryVM.AmountDelivery = DeliverySelect.AmountDelivery;
                await displayRootRegistry.ShowModalPresentation(deliveryVM);
                if (deliveryVM.PressOk)
                {
                    if (deliveryVM.OrderNumber != "")
                    {
                        DeliverySelect.OrderNumber = deliveryVM.OrderNumber;
                        DeliverySelect.Color = "Blue";
                        db.Entry(DeliverySelect).State = EntityState.Modified;
                        db.SaveChanges();

                        Deliveries = null;
                        LoadDeliveriesDB(CompanyName, false);
                    }
                    else
                    {
                        DeliverySelect.OrderNumber = deliveryVM.OrderNumber;
                        DeliverySelect.Color = "Red";
                        db.Entry(DeliverySelect).State = EntityState.Modified;
                        db.SaveChanges();

                        Deliveries = null;
                        LoadDeliveriesDB(CompanyName, false);
                    }
                }
            }
        }));

        //**********************************
        public Command AddInvoice => _addInvoice ?? (_addInvoice = new Command(async obj =>
        {
            var displayRootRegistry = (Application.Current as App).displayRootRegistry;
            var invoiceViewModel = new InvoiceViewModel();
            invoiceViewModel.Quota = Quotations;
            await displayRootRegistry.ShowModalPresentation(invoiceViewModel);
            if (invoiceViewModel.PressOk)
            {
                var temp = Clients.FirstOrDefault(q => q.Id == invoiceViewModel.QuotaSelect.ClientId); //db.Clients.FirstOrDefault(q => q.Id == invoiceViewModel.QuotaSelect.ClientId);

                Invoice inv = new Invoice()
                {
                    DateInvoice = DateTime.Today,
                    NumberQuota = invoiceViewModel.QuotaSelect.NumberQuota,
                    OrderNumber = invoiceViewModel.OrderNumber,
                    UpNumber = invoiceViewModel.UpNumber,
                    QuotaId = invoiceViewModel.QuotaSelect.Id,
                    FirstName = temp.PrimaryFirstName,
                    LastName = temp.PrimaryLastName,
                    PhoneNumber = temp.PrimaryPhoneNumber,
                    Email = temp.PrimaryEmail,
                    CompanyName = invoiceViewModel.QuotaSelect.CompanyName
                };

                var deliveryToArchive = db.Deliveries.Where(d => d.QuotaId == inv.QuotaId);
                foreach (var item in deliveryToArchive)
                {
                    item.IsArchive = true;
                    db.Entry(item).State = EntityState.Modified;
                }
                Deliveries = null;
                LoadDeliveriesDB(CompanyName, false);
                IsCheckedArchiveDelivery = false;

                db.Invoices.Add(inv);
                db.SaveChanges();
                //Invoices = null;
                //Invoices = db.Invoices.Local.ToBindingList();

                //int? NewInvoiceId = db.Invoices.Local.ToBindingList().OrderByDescending(q => q.Id).FirstOrDefault().Id;
                //AddMaterialProfit(NewInvoiceId);
                //MaterialProfitCalculate(MaterialProfitSelect);
                //AddLabourProfit(NewInvoiceId);
                //LabourProfitCalculate(LabourProfitSelect);
                //AddDebtsTab(NewInvoiceId);
                //InvoiceCalculate(NewInvoiceId);
                int? NewInvoiceId = Invoices.OrderByDescending(q => q.Id).FirstOrDefault().Id;
                AddMaterialProfit(NewInvoiceId);
                AddLabourProfit(NewInvoiceId);
                AddDebtsTab(NewInvoiceId);
                var material = MaterialProfits.FirstOrDefault(m => m.InvoiceId == NewInvoiceId);
                var labour = LabourProfits.FirstOrDefault(l => l.InvoiceId == NewInvoiceId);
                MaterialProfitCalculate(material);
                LabourProfitCalculate(labour);
                InvoiceCalculate(NewInvoiceId);
            }
        }));
        public Command InsInvoice => _insInvoice ?? (_insInvoice = new Command(async obj =>
        {
            if (InvoiceSelect != null)
            {
                var displayRootRegistry = (Application.Current as App).displayRootRegistry;
                var invoiceViewModel = new InvoiceViewModel();
                invoiceViewModel.Quota = Quotations;
                invoiceViewModel.QuotaSelect = db.Quotations.FirstOrDefault(q => q.Id == InvoiceSelect.QuotaId);
                invoiceViewModel.OrderNumber = InvoiceSelect.OrderNumber;
                invoiceViewModel.UpNumber = InvoiceSelect.UpNumber;
                await displayRootRegistry.ShowModalPresentation(invoiceViewModel);
                if (invoiceViewModel.PressOk)
                {
                    //var temp = db.Clients.FirstOrDefault(q => q.Id == invoiceViewModel.QuotaSelect.ClientId);   // Нахрена змінювати Квоту ???

                    //InvoiceSelect.QuotaId = invoiceViewModel.QuotaSelect.Id;
                    InvoiceSelect.OrderNumber = invoiceViewModel.OrderNumber;
                    InvoiceSelect.UpNumber = invoiceViewModel.UpNumber;
                    //InvoiceSelect.FulNameClient = temp.PrimaryFirstName + " " + temp.PrimaryLastName;

                    db.Entry(InvoiceSelect).State = EntityState.Modified;
                    db.SaveChanges();

                    Invoices = null;
                    LoadInvoicesDB(CompanyName);
                }
            }
        }));
        public Command DelInvoice => _delInvoice ?? (_delInvoice = new Command(obj =>
        {
            if (InvoiceSelect != null)
            {
                var materialProfit = MaterialProfits.FirstOrDefault(m => m.InvoiceId == InvoiceSelect.Id); //db.MaterialProfits.FirstOrDefault(m => m.InvoiceId == InvoiceSelect.Id);
                var material = db.Materials.Where(mat => mat.MaterialProfitId == materialProfit.Id);
                var labourPrifit = LabourProfits.FirstOrDefault(l => l.InvoiceId == InvoiceSelect.Id); //db.LabourProfits.FirstOrDefault(l => l.InvoiceId == InvoiceSelect.Id);
                var labour = db.Labours.Where(l => l.LabourProfitId == labourPrifit.Id);
                var labourCon = db.LabourContractors.Where(c => c.LabourProfitId == labourPrifit.Id);
                var debts = Debts.Where(d => d.InvoiceId == InvoiceSelect.Id); //db.Debts.Where(d => d.InvoiceId == InvoiceSelect.Id);
                db.LabourContractors.RemoveRange(labourCon);
                db.Labours.RemoveRange(labour);
                db.LabourProfits.Remove(labourPrifit);
                db.Materials.RemoveRange(material);
                db.MaterialProfits.Remove(materialProfit);
                db.Debts.RemoveRange(debts);
                db.Invoices.Remove(InvoiceSelect);     // Обовязково треба звертати увагу на залежність !!! ЧЕРГА !!!                
                db.SaveChanges();
                Invoices = null;
                LoadInvoicesDB(CompanyName);
                MaterialProfits = null;
                LoadMaterialProfitsDB(CompanyName);
                LabourProfits = null;
                LoadLabourProfitsDB(CompanyName);
                Debts = null;
                LoadDebtsDB(CompanyName);
            }
        }));
        public Command PrintInvoice => _printInvoice ?? (_printInvoice = new Command(obj =>
        {
            try
            {

                if (CompanyName == "CMO")
                {
                    InvoicesPrintToExcelCMO("\\Blanks\\InvoicePDF");   //  "\\Blanks\\InvoicePDF.xltm"
                }
                else
                {
                    InvoicePrintToExcelNL("\\Blanks\\NLInvoicePDF");   //  "\\Blanks\\NLInvoicePDF.xltm"
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error message: " + Environment.NewLine +
                                       ex.Message + Environment.NewLine + Environment.NewLine +
                                       "StackTrace message: " + Environment.NewLine +
                                       ex.StackTrace, "Warning !!!");
            }
        }));
        public Command SearchInvoice => _searchInvoice ?? (_searchInvoice = new Command(obj =>
        {
            string search = obj.ToString();
            if (search == "")
            {
                LoadInvoicesDB(CompanyName);
            }
            else
            {
                Invoices = Invoices.Where(n => n.CompanyName == CompanyName && n.FullSearch.ToUpper().Contains(search.ToUpper()));
            }
        }));

        //**********************************
        public Command AddWorkOrder => _addWorkOrder ?? (_addWorkOrder = new Command(async obj =>
        {
            var displayRootRegistry = (Application.Current as App).displayRootRegistry;
            var work = new WorkOrderViewModel(ref db, EnumClient.Add, null);
            work.Quotations = Quotations;
            await displayRootRegistry.ShowModalPresentation(work);
            LoadWorkOrdersDB(CompanyName);
        }));
        public Command InsWorkOrder => _insWorkOrder ?? (_insWorkOrder = new Command(async obj =>
        {
            if (WorkOrderSelect != null)
            {
                var displayRootRegistry = (Application.Current as App).displayRootRegistry;
                var work = new WorkOrderViewModel(ref db, EnumClient.Ins, WorkOrderSelect.Id);
                work.Quotations = Quotations;
                await displayRootRegistry.ShowModalPresentation(work);

                WorkOrderSelect.Parking = work.Parking;
                WorkOrderSelect.DateServices = work.ServiceDate;
                WorkOrderSelect.DateCompletion = work.CompletionDate;
                //WorkOrderSelect.Trim = work.TrimSelect;
                //WorkOrderSelect.Colour = work.ColourSelect;
                WorkOrderSelect.Notes = work.Notes;
                //WorkOrderSelect.Baseboard = work.BaseboardSelect;
                //WorkOrderSelect.ReplacingYesNo = work.ReplacingSelect;
                //WorkOrderSelect.ReplacingQuantity = work.Pieces;

                db.Entry(WorkOrderSelect).State = EntityState.Modified;
                db.SaveChanges();

                WorkOrders = null;
                LoadWorkOrdersDB(CompanyName);
            }
        }));
        public Command DelWorkOrder => _delWorkOrder ?? (_delWorkOrder = new Command(obj =>
        {
            if (WorkOrderSelect != null)
            {
                var work = db.WorkOrder_Works.Where(w => w.WorkOrderId == WorkOrderSelect.Id);
                var accessories = db.WorkOrder_Accessories.Where(a => a.WorkOrderId == WorkOrderSelect.Id);
                var install = db.WorkOrder_Installations.Where(i => i.WorkOrderId == WorkOrderSelect.Id);
                var contract = db.WorkOrder_Contractors.Where(c => c.WorkOrderId == WorkOrderSelect.Id);
                db.WorkOrder_Contractors.RemoveRange(contract);
                db.WorkOrder_Installations.RemoveRange(install);
                db.WorkOrder_Accessories.RemoveRange(accessories);
                db.WorkOrder_Works.RemoveRange(work);
                db.WorkOrders.Remove(WorkOrderSelect);
                db.SaveChanges();
                WorkOrders = null;
                LoadWorkOrdersDB(CompanyName);
            }
        }));
        public Command PrintWorkOrder => _printWorkOrder ?? (_printWorkOrder = new Command(obj =>
        {
            if (WorkOrderSelect != null)
            {
                try
                {
                    var contractors = db.WorkOrder_Contractors.Where(c => c.WorkOrderId == WorkOrderSelect.Id).OrderBy(c => c.Contractor);
                    foreach (var contractor in contractors)
                    {
                        int? room = db.WorkOrder_Works.Where(w => w.Contractor == contractor.Contractor)?.Count();
                        int? accessories = db.WorkOrder_Accessories.Where(a => a.Contractor == contractor.Contractor)?.Count();
                        if (room > 4 || accessories > 7)
                        {
                            WorkOrderPrintMaxToExcel(contractor, "\\Blanks\\WorkOrderPDFmax");   //   "\\Blanks\\WorkOrderPDFmax.xltm"
                        }
                        else
                        {
                            WorkOrderPrintMinToExcel(contractor, "\\Blanks\\WorkOrderPDFmin");   //   "\\Blanks\\WorkOrderPDFmin.xltm"
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error message: " + Environment.NewLine +
                                           ex.Message + Environment.NewLine + Environment.NewLine +
                                           "StackTrace message: " + Environment.NewLine +
                                           ex.StackTrace, "Warning !!!");
                }
            }
        }));
        public Command SearchWorkOrder => _searchWorkOrder ?? (_searchWorkOrder = new Command(obj =>
        {
            string search = obj.ToString();
            if (search == "")
            {
                LoadWorkOrdersDB(CompanyName);
            }
            else
            {
                WorkOrders = WorkOrders.Where(n => n.CompanyName == CompanyName && n.FullSearch.ToUpper().Contains(search.ToUpper()));
            }
        }));
        //**********************************
        public Command InsMaterialProfit => _insMaterialProfit ?? (_insMaterialProfit = new Command(async obj =>
        {
            if (MaterialProfitSelect != null)
            {
                int? temp = MaterialProfitSelect.InvoiceId;
                var displayRootRegistry = (Application.Current as App).displayRootRegistry;
                var mp = new MaterialProfitViewModel(ref db, MaterialProfitSelect);
                await displayRootRegistry.ShowModalPresentation(mp);
                MaterialProfitCalculate(MaterialProfitSelect);
                InvoiceCalculate(temp);
            }
        }));
        public Command PrintMaterialProfit => _printMaterialProfit ?? (_printMaterialProfit = new Command(obj =>
        {
            if (MaterialProfitSelect != null)
            {
                try
                {
                    int i = 0;
                    var invoice = db.Invoices.FirstOrDefault(inv => inv.Id == MaterialProfitSelect.InvoiceId);
                    var quota = db.Quotations.FirstOrDefault(q => q.Id == invoice.QuotaId);
                    var clientData = db.Clients.FirstOrDefault(c => c.Id == quota.ClientId);
                    var material = db.Materials.Where(m => m.MaterialProfitId == MaterialProfitSelect.Id);

                    Excel.Application ExcelApp = new Excel.Application();
                    Excel.Workbook ExcelWorkBook;
                    ExcelWorkBook = ExcelApp.Workbooks.Open(Environment.CurrentDirectory + "\\Blanks\\MaterialProfitPDF");   //Вказуємо шлях до шаблону   "\\Blanks\\MaterialProfitPDF.xltm"

                    ExcelApp.Cells[2, 2] = invoice.NumberInvoice;
                    ExcelApp.Cells[4, 2] = clientData.PrimaryFirstName + " " + clientData.PrimaryLastName;
                    ExcelApp.Cells[5, 2] = clientData.SecondaryFirstName + " " + clientData.SecondaryLastName;
                    ExcelApp.Cells[6, 2] = clientData.CompanyName;
                    ExcelApp.Cells[2, 11] = MaterialProfitSelect.InvoiceDate;
                    ExcelApp.Cells[4, 8] = quota.JobDescription;
                    ExcelApp.Cells[5, 8] = clientData.AddressSiteStreet + ", " + clientData.AddressSiteCity + ", " + clientData.AddressSiteProvince + ", " + clientData.AddressSitePostalCode + ", " + clientData.AddressSiteCountry;


                    i = 11;
                    foreach (var item in material)
                    {
                        if (item.Groupe == "FLOORING")
                        {
                            ExcelApp.Cells[i, 1] = item.Item;
                            ExcelApp.Cells[i, 2] = item.Description;
                            ExcelApp.Cells[i, 4] = item.Quantity;
                            ExcelApp.Cells[i, 5] = item.Rate;
                            ExcelApp.Cells[i, 6] = item.Price;
                            ExcelApp.Cells[i, 7] = (item.CostQuantity != 0) ? (item.CostQuantity) : 0m;
                            ExcelApp.Cells[i, 8] = (item.CostUnitPrice != 0) ? (item.CostUnitPrice) : 0m;
                            ExcelApp.Cells[i, 9] = (item.CostSubtotal != 0) ? (item.CostSubtotal) : 0m;
                            ExcelApp.Cells[i, 10] = (item.CostEPRate != 0) ? (item.CostEPRate) : 0m;
                            ExcelApp.Cells[i, 11] = (item.CostTax != 0) ? (item.CostTax) : 0m;
                            ExcelApp.Cells[i, 12] = (item.CostTotal != 0) ? (item.CostTotal) : 0m;
                            i++;
                        }
                    }

                    i = 20;
                    foreach (var item in material)
                    {
                        if (item.Groupe == "ACCESSORIES")
                        {
                            ExcelApp.Cells[i, 1] = item.Item;
                            ExcelApp.Cells[i, 2] = item.Description;
                            ExcelApp.Cells[i, 4] = item.Quantity;
                            ExcelApp.Cells[i, 5] = item.Rate;
                            ExcelApp.Cells[i, 6] = item.Price;
                            ExcelApp.Cells[i, 7] = (item.CostQuantity != 0) ? (item.CostQuantity) : 0m;
                            ExcelApp.Cells[i, 8] = (item.CostUnitPrice != 0) ? (item.CostUnitPrice) : 0m;
                            ExcelApp.Cells[i, 9] = (item.CostSubtotal != 0) ? (item.CostSubtotal) : 0m;
                            ExcelApp.Cells[i, 10] = (item.CostEPRate != 0) ? (item.CostEPRate) : 0m;
                            ExcelApp.Cells[i, 11] = (item.CostTax != 0) ? (item.CostTax) : 0m;
                            ExcelApp.Cells[i, 12] = (item.CostTotal != 0) ? (item.CostTotal) : 0m;
                            i++;
                        }
                    }

                    i = 42;
                    foreach (var item in material)
                    {
                        if (item.Groupe == "OTHER")
                        {
                            ExcelApp.Cells[i, 1] = item.Item;
                            ExcelApp.Cells[i, 2] = item.Description;
                            ExcelApp.Cells[i, 3] = item.CostQuantity;
                            ExcelApp.Cells[i, 4] = item.CostUnitPrice;
                            ExcelApp.Cells[i, 5] = item.CostSubtotal;
                            ExcelApp.Cells[i, 6] = item.CostEPRate;
                            ExcelApp.Cells[i, 7] = item.CostTax;
                            ExcelApp.Cells[i, 8] = item.CostTotal;
                            i++;
                        }
                    }



                    ExcelApp.Cells[36, 4] = MaterialProfitSelect.MaterialSubtotal;
                    ExcelApp.Cells[37, 4] = MaterialProfitSelect.MaterialTax;
                    ExcelApp.Cells[38, 4] = MaterialProfitSelect.MaterialTotal;

                    ExcelApp.Cells[55, 7] = MaterialProfitSelect.CostMaterialSubtotal;
                    ExcelApp.Cells[56, 7] = MaterialProfitSelect.CostMaterialTax;
                    ExcelApp.Cells[57, 7] = MaterialProfitSelect.CostMaterialTotal;

                    ExcelApp.Cells[61, 5] = MaterialProfitSelect.ProfitBeforTax;
                    ExcelApp.Cells[62, 5] = MaterialProfitSelect.ProfitTax;
                    ExcelApp.Cells[63, 5] = MaterialProfitSelect.ProfitInclTax;
                    ExcelApp.Cells[64, 5] = MaterialProfitSelect.ProfitDiscount;
                    ExcelApp.Cells[65, 5] = MaterialProfitSelect.ProfitTotal;



                    ExcelApp.Calculate();
                    ExcelApp.Cells[1, 14] = "1";   // Записуємо дані в .pdf    
                    ExcelApp.Calculate();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error message: " + Environment.NewLine +
                                           ex.Message + Environment.NewLine + Environment.NewLine +
                                           "StackTrace message: " + Environment.NewLine +
                                           ex.StackTrace, "Warning !!!");
                }

            }
        }));
        public Command SearchMaterialProfit => _searchMaterialProfit ?? (_searchMaterialProfit = new Command(obj =>
        {
            string search = obj.ToString();
            if (search == "")
            {
                LoadMaterialProfitsDB(CompanyName);
            }
            else
            {
                MaterialProfits = MaterialProfits.Where(n => n.CompanyName == CompanyName && n.FullSearch.ToUpper().Contains(search.ToUpper()));
            }
        }));
        public Command ClickMaterialProfit => _clickMaterialProfit ?? (_clickMaterialProfit = new Command(obj =>
        {
            if (MaterialProfitSelect != null)
            {
                MaterialProfitSelect.Companion = (MaterialProfitSelect.Companion) ? (false) : (true);
                db.Entry(MaterialProfitSelect).State = EntityState.Modified;
                db.SaveChanges();

                MaterialProfits = null;
                LoadMaterialProfitsDB(CompanyName);
            }
        }));
        public Command LoadFotoMaterialProfit => _loadFotoMaterialProfit ?? (_loadFotoMaterialProfit = new Command(obj =>
        {
            if (MaterialProfitSelect != null)
            {
                string[] files = OpenFileArray("Image files (*.png;*.jpeg;*.jpg;*.JPG)|*.png;*.jpeg;*.jpg;*.JPG");

                if (worker != null && worker.IsBusy) { return; }
                worker = new BackgroundWorker();
                worker.DoWork += worker_DoWork;
                worker.RunWorkerAsync();
                void worker_DoWork(object sender, DoWorkEventArgs e)
                {
                    ProgressStart();

                    // Основні затратні задачі
                    string SecondDir = Directory.GetCurrentDirectory() + "\\Foto";

                    if (files != null)
                    {
                        foreach (string fil in files)
                        {
                            File.Copy(fil, Path.Combine(SecondDir, MaterialProfitSelect.Id + "_mat_" + Path.GetRandomFileName() + Path.GetExtension(fil)), true);
                        }
                    }
                    //************************

                    ProgressStop();
                }

            }
        }));
        public Command ViewFotoMaterialProfit => _viewFotoMaterialProfit ?? (_viewFotoMaterialProfit = new Command(async obj =>
        {
            var displayRootRegistry = (Application.Current as App).displayRootRegistry;
            var foto = new FotoViewModel(MaterialProfitSelect.Id, "_mat_*");
            await displayRootRegistry.ShowModalPresentation(foto);
        }));
        //**********************************
        public Command InsLabourProfit => _insLabourProfit ?? (_insLabourProfit = new Command(async obj =>
        {
            if (LabourProfitSelect != null)
            {
                int? temp = LabourProfitSelect.InvoiceId;
                var displayRootRegistry = (Application.Current as App).displayRootRegistry;
                var lp = new LabourProfitViewModel(ref db, LabourProfitSelect);
                await displayRootRegistry.ShowModalPresentation(lp);
                InsDebtsTab(LabourProfitSelect.InvoiceId);
                LabourProfitCalculate(LabourProfitSelect);
                InvoiceCalculate(temp);
            }
        }));
        public Command PrintLabourProfit => _printLabourProfit ?? (_printLabourProfit = new Command(obj =>
        {
            if (LabourProfitSelect != null)
            {
                try
                {

                    int i = 0;
                    var invoice = db.Invoices.FirstOrDefault(inv => inv.Id == LabourProfitSelect.InvoiceId);
                    var quota = db.Quotations.FirstOrDefault(q => q.Id == invoice.QuotaId);
                    var clientData = db.Clients.FirstOrDefault(c => c.Id == quota.ClientId);
                    var labour = db.Labours.Where(m => m.LabourProfitId == LabourProfitSelect.Id);
                    var contractor = db.LabourContractors.Where(c => c.LabourProfitId == LabourProfitSelect.Id);

                    Excel.Application ExcelApp = new Excel.Application();
                    Excel.Workbook ExcelWorkBook;
                    ExcelWorkBook = ExcelApp.Workbooks.Open(Environment.CurrentDirectory + "\\Blanks\\LabourProfitPDF");   //Вказуємо шлях до шаблону   "\\Blanks\\LabourProfitPDF.xltm"

                    ExcelApp.Cells[3, 2] = invoice.NumberInvoice;
                    ExcelApp.Cells[4, 2] = clientData.PrimaryFirstName + " " + clientData.PrimaryLastName;
                    ExcelApp.Cells[5, 2] = clientData.SecondaryFirstName + " " + clientData.SecondaryLastName;
                    ExcelApp.Cells[6, 2] = clientData.CompanyName;
                    ExcelApp.Cells[3, 5] = invoice.DateInvoice;
                    ExcelApp.Cells[4, 5] = quota.JobDescription;
                    ExcelApp.Cells[5, 5] = clientData.AddressSiteStreet + ", " + clientData.AddressSiteCity + ", " + clientData.AddressSiteProvince + ", " + clientData.AddressSitePostalCode + ", " + clientData.AddressSiteCountry;


                    i = 10;
                    foreach (var item in labour)
                    {
                        if (item.Groupe == "INSTALLATION")
                        {
                            ExcelApp.Cells[i, 1] = item.Item;
                            ExcelApp.Cells[i, 2] = item.Description;
                            ExcelApp.Cells[i, 3] = item.Contractor;
                            ExcelApp.Cells[i, 4] = item.Quantity;
                            ExcelApp.Cells[i, 5] = item.Rate;
                            ExcelApp.Cells[i, 6] = item.Price;
                            ExcelApp.Cells[i, 7] = item.Percent;
                            ExcelApp.Cells[i, 8] = item.Payout;
                            ExcelApp.Cells[i, 9] = item.Profit;

                            i++;
                        }
                    }

                    i = 19;
                    foreach (var item in labour)
                    {
                        if (item.Groupe == "DEMOLITION")
                        {
                            ExcelApp.Cells[i, 1] = item.Item;
                            ExcelApp.Cells[i, 2] = item.Description;
                            ExcelApp.Cells[i, 3] = item.Contractor;
                            ExcelApp.Cells[i, 4] = item.Quantity;
                            ExcelApp.Cells[i, 5] = item.Rate;
                            ExcelApp.Cells[i, 6] = item.Price;
                            ExcelApp.Cells[i, 7] = item.Percent;
                            ExcelApp.Cells[i, 8] = item.Payout;
                            ExcelApp.Cells[i, 9] = item.Profit;

                            i++;
                        }
                    }

                    i = 25;
                    foreach (var item in labour)
                    {
                        if (item.Groupe == "OPTIONAL SERVICES")
                        {
                            ExcelApp.Cells[i, 1] = item.Item;
                            ExcelApp.Cells[i, 2] = item.Description;
                            ExcelApp.Cells[i, 3] = item.Contractor;
                            ExcelApp.Cells[i, 4] = item.Quantity;
                            ExcelApp.Cells[i, 5] = item.Rate;
                            ExcelApp.Cells[i, 6] = item.Price;
                            ExcelApp.Cells[i, 7] = item.Percent;
                            ExcelApp.Cells[i, 8] = item.Payout;
                            ExcelApp.Cells[i, 9] = item.Profit;

                            i++;
                        }
                    }

                    i = 33;
                    foreach (var item in labour)
                    {
                        if (item.Groupe == "FLOORING DELIVERY")
                        {
                            ExcelApp.Cells[i, 1] = item.Item;
                            ExcelApp.Cells[i, 2] = item.Description;
                            ExcelApp.Cells[i, 3] = item.Contractor;
                            ExcelApp.Cells[i, 4] = item.Quantity;
                            ExcelApp.Cells[i, 5] = item.Rate;
                            ExcelApp.Cells[i, 6] = item.Price;
                            ExcelApp.Cells[i, 7] = item.Percent;
                            ExcelApp.Cells[i, 8] = item.Payout;
                            ExcelApp.Cells[i, 9] = item.Profit;

                            i++;
                        }
                    }


                    ExcelApp.Cells[42, 3] = LabourProfitSelect.CollectedSubtotal;
                    ExcelApp.Cells[43, 3] = LabourProfitSelect.CollectedGST;
                    ExcelApp.Cells[44, 3] = LabourProfitSelect.CollectedTotal;

                    ExcelApp.Cells[42, 4] = LabourProfitSelect.PayoutSubtotal;
                    ExcelApp.Cells[43, 4] = LabourProfitSelect.PayoutGST;
                    ExcelApp.Cells[44, 4] = LabourProfitSelect.PayoutTotal;

                    ExcelApp.Cells[42, 5] = LabourProfitSelect.StoreSubtotal;
                    ExcelApp.Cells[43, 5] = LabourProfitSelect.StoreGST;
                    ExcelApp.Cells[44, 5] = LabourProfitSelect.StoreTotal;

                    ExcelApp.Cells[45, 5] = LabourProfitSelect.Discount;
                    ExcelApp.Cells[46, 5] = LabourProfitSelect.ProfitTotal;

                    i = 50;
                    foreach (var item in contractor)
                    {
                        ExcelApp.Cells[i, 2] = item.Contractor;
                        ExcelApp.Cells[i, 3] = item.Payout;
                        ExcelApp.Cells[i, 4] = item.Adjust;
                        ExcelApp.Cells[i, 5] = item.Total;
                        ExcelApp.Cells[i, 6] = item.TAX;
                        ExcelApp.Cells[i, 7] = item.GST;
                        ExcelApp.Cells[i, 8] = item.TotalContractor;

                        i++;
                    }

                    ExcelApp.Calculate();
                    ExcelApp.Cells[1, 11] = "1";   // Записуємо дані в .pdf    
                    ExcelApp.Calculate();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error message: " + Environment.NewLine +
                                           ex.Message + Environment.NewLine + Environment.NewLine +
                                           "StackTrace message: " + Environment.NewLine +
                                           ex.StackTrace, "Warning !!!");
                }

            }
        }));
        public Command SearchLabourProfit => _searchLabourProfit ?? (_searchLabourProfit = new Command(obj =>
        {
            string search = obj.ToString();
            if (search == "")
            {
                LoadLabourProfitsDB(CompanyName);
            }
            else
            {
                LabourProfits = LabourProfits.Where(n => n.CompanyName == CompanyName && n.FullSearch.ToUpper().Contains(search.ToUpper()));
            }
        }));
        public Command ClickLabourPrifit => _clickLabourProfit ?? (_clickLabourProfit = new Command(obj =>
        {
            if (LabourProfitSelect != null)
            {
                LabourProfitSelect.Companion = (LabourProfitSelect.Companion) ? (false) : (true);
                db.Entry(LabourProfitSelect).State = EntityState.Modified;
                db.SaveChanges();

                LabourProfits = null;
                LoadLabourProfitsDB(CompanyName);
            }
        }));
        //**********************************
        public Command AddDebts => _addDebts ?? (_addDebts = new Command(async obj =>
        {
            var displayRootRegistry = (Application.Current as App).displayRootRegistry;
            var addDebts = new DebtsViewModel(db, DebtSelect);

            await displayRootRegistry.ShowModalPresentation(addDebts);

            if (addDebts.PressOk)
            {
                Debts debts = new Debts()
                {
                    InvoiceId = addDebts.InvoiceSelect.Id,
                    InvoiceDate = addDebts.InvoiceSelect.DateInvoice,
                    InvoiceNumber = addDebts.InvoiceSelect.NumberInvoice,
                    FirstName = addDebts.InvoiceSelect.FirstName,
                    LastName = addDebts.InvoiceSelect.LastName,
                    Email = addDebts.InvoiceSelect.Email,
                    PhoneNumber = addDebts.InvoiceSelect.PhoneNumber,
                    NameDebts = addDebts.NameDebts,
                    DescriptionDebts = addDebts.Description,
                    AmountDebts = decimal.Round(addDebts.Amount, 2),
                    ReadOnly = false
                };
                db.Debts.Add(debts);
                db.SaveChanges();
                Debts = null;
                LoadDebtsDB(CompanyName);
            }

        }));
        public Command InsDebts => _insDebts ?? (_insDebts = new Command(async obj =>
        {
            if (DebtSelect != null && DebtSelect.ReadOnly == false)
            {
                var displayRootRegistry = (Application.Current as App).displayRootRegistry;
                var insDebts = new DebtsViewModel(db, DebtSelect)
                {
                    NameDebts = DebtSelect.NameDebts,
                    Description = DebtSelect.DescriptionDebts,
                    Amount = DebtSelect.AmountDebts
                };
                await displayRootRegistry.ShowModalPresentation(insDebts);

                if (insDebts.PressOk)
                {

                    DebtSelect.InvoiceId = insDebts.InvoiceSelect.Id;
                    DebtSelect.InvoiceDate = insDebts.InvoiceSelect.DateInvoice;
                    DebtSelect.InvoiceNumber = insDebts.InvoiceSelect.NumberInvoice;
                    DebtSelect.FirstName = insDebts.InvoiceSelect.FirstName;
                    DebtSelect.LastName = insDebts.InvoiceSelect.LastName;
                    DebtSelect.Email = insDebts.InvoiceSelect.Email;
                    DebtSelect.PhoneNumber = insDebts.InvoiceSelect.PhoneNumber;
                    DebtSelect.NameDebts = insDebts.NameDebts;
                    DebtSelect.DescriptionDebts = insDebts.Description;
                    DebtSelect.AmountDebts = decimal.Round(insDebts.Amount, 2);
                    DebtSelect.ReadOnly = false;

                    db.Entry(DebtSelect).State = EntityState.Modified;
                    db.SaveChanges();
                    Debts = null;
                    LoadDebtsDB(CompanyName);
                }
            }
        }));
        public Command DelDebts => _delDebts ?? (_delDebts = new Command(obj =>
        {
            if (DebtSelect != null && DebtSelect.ReadOnly == false)
            {
                db.Debts.Remove(DebtSelect);
                db.SaveChanges();
                Debts = null;
                LoadDebtsDB(CompanyName);
            }
        }));
        public Command PaymentDebts => _paymentDebts ?? (_paymentDebts = new Command(async obj =>
        {
            var displayRootRegistry = (Application.Current as App).displayRootRegistry;
            var PayDebts = new DebtsPaymentViewModel();
            if (DebtSelect.DatePayment > DateTime.MinValue)
            {
                PayDebts.Date = DebtSelect.DatePayment;
            }
            if (DebtSelect.AmountPayment != 0m)
            {
                PayDebts.Amount = DebtSelect.AmountPayment;
            }
            else
            {
                PayDebts.Amount = DebtSelect.AmountDebts;
            }
            PayDebts.Description = DebtSelect.DescriptionPayment;

            await displayRootRegistry.ShowModalPresentation(PayDebts);
            if (PayDebts.PressOk)
            {
                DebtSelect.DatePayment = PayDebts.Date;
                DebtSelect.DescriptionPayment = PayDebts.Description;
                DebtSelect.AmountPayment = decimal.Round(PayDebts.Amount, 2);
                if (PayDebts.Amount != 0m)
                {
                    DebtSelect.ColorPayment = "Silver";
                    DebtSelect.Payment = true;
                }
                else
                {
                    DebtSelect.ColorPayment = "Red";
                    DebtSelect.Payment = false;
                }

                db.Entry(DebtSelect).State = EntityState.Modified;
                db.SaveChanges();
                Debts = null;
                LoadDebtsDB(CompanyName);
            }

        }));
        public Command SearchDebts => _searchDebts ?? (_searchDebts = new Command(obj =>
        {
            string search = obj.ToString();
            if (search == "")
            {
                LoadDebtsDB(CompanyName);
            }
            else
            {
                Debts = Debts.Where(n => n.CompanyName == CompanyName && n.FullSearch.ToUpper().Contains(search.ToUpper()));
            }
        }));
        //**********************************

        public Command LoadReport => _loadReport ?? (_loadReport = new Command(obj =>
        {

            if (IsCheckedProfit)
            {
                IsVisibleDateSelect = Visibility.Visible;
                IsVisibleDebtsReport = Visibility.Collapsed;
                IsVisibleTotalReport = Visibility.Collapsed;
                IsVisiblePayrollReport = Visibility.Collapsed;
                IsVisibleAmountReport = Visibility.Collapsed;
                IsVisibleProfitReport = Visibility.Visible;
                IsVisibleExpensesReport = Visibility.Collapsed;
                MaterialReport = ReportMaterial(ReportDateFrom, ReportDateTo);
                LabourReport = ReportLabour(ReportDateFrom, ReportDateTo);
                CountReport = "Total for the selected period - " + MaterialReport?.InvoiceNumber;
            }
            else if (IsCheckedPayroll)
            {
                IsVisibleDateSelect = Visibility.Visible;
                IsVisiblePayrollReport = Visibility.Visible;
                IsVisibleProfitReport = Visibility.Collapsed;
                IsVisibleTotalReport = Visibility.Collapsed;
                IsVisibleAmountReport = Visibility.Collapsed;
                IsVisibleDebtsReport = Visibility.Collapsed;
                IsVisibleExpensesReport = Visibility.Collapsed;
                ReportPayrollToWorks = ReportPayrollWork(ReportDateFrom, ReportDateTo);
                ReportPayrolls = ReportPayroll(ReportPayrollToWorks);
            }
            else if (IsCheckedAmount)
            {
                IsVisibleDateSelect = Visibility.Visible;
                IsVisibleAmountReport = Visibility.Visible;
                IsVisiblePayrollReport = Visibility.Collapsed;
                IsVisibleProfitReport = Visibility.Collapsed;
                IsVisibleTotalReport = Visibility.Collapsed;
                IsVisibleDebtsReport = Visibility.Collapsed;
                IsVisibleExpensesReport = Visibility.Collapsed;
                ReportAmounts = ReportAmountGet(ReportDateFrom, ReportDateTo);
            }
            else if (IsCheckedDebts)
            {
                IsVisibleDateSelect = Visibility.Visible;
                IsVisibleDebtsReport = Visibility.Visible;
                IsVisibleTotalReport = Visibility.Collapsed;
                IsVisiblePayrollReport = Visibility.Collapsed;
                IsVisibleAmountReport = Visibility.Collapsed;
                IsVisibleProfitReport = Visibility.Collapsed;
                IsVisibleExpensesReport = Visibility.Collapsed;
                AmountDebtsReport = ReportAmountDebts(ReportDateFrom, ReportDateTo);
                PaymentDebtsReport = ReportPaymentDebts(ReportDateFrom, ReportDateTo);
                LabelAmountDebts = decimal.Round(AmountDebtsReport.Select(a => a.AmountDebts)?.Sum() ?? 0m, 2);
                LabelPaymentDebts = decimal.Round(PaymentDebtsReport.Select(p => p.AmountPayment)?.Sum() ?? 0m, 2);
            }
            else if (IsCheckedExpenses)
            {
                IsVisibleExpensesReport = Visibility.Visible;
                IsVisibleDateSelect = Visibility.Collapsed;
                IsVisibleAmountReport = Visibility.Collapsed;
                IsVisiblePayrollReport = Visibility.Collapsed;
                IsVisibleProfitReport = Visibility.Collapsed;
                IsVisibleTotalReport = Visibility.Collapsed;
                IsVisibleDebtsReport = Visibility.Collapsed;
                ActiveExpenses = ReportActivExpenses(DateYearSelect, DateMonthSelect);
                PaymentExpenses = ReportPaymentExpenses(DateYearSelect, DateMonthSelect);
                LabelActiveExpenses = decimal.Round(ActiveExpenses.Select(e => e.Amounts)?.Sum() ?? 0m, 2);
                LabelPaymentExpenses = decimal.Round(PaymentExpenses.Select(e => e.AmountPaid)?.Sum() ?? 0m, 2);
            }
            else
            {
                IsVisibleDateSelect = Visibility.Visible;
                IsVisibleTotalReport = Visibility.Visible;
                IsVisibleProfitReport = Visibility.Collapsed;
                IsVisiblePayrollReport = Visibility.Collapsed;
                IsVisibleAmountReport = Visibility.Collapsed;
                IsVisibleDebtsReport = Visibility.Collapsed;
                IsVisibleExpensesReport = Visibility.Collapsed;
                TotalReport = ReportTotalGrid(ReportDateFrom, ReportDateTo);
                TotalReportForLabel = ReportTotalLabel(ReportDateFrom, ReportDateTo, TotalReport);
            }
        }));
        public Command LoadExpenses => _loadExpenses ?? (_loadExpenses = new Command(async obj =>
        {
            var displayRootRegistry = (Application.Current as App).displayRootRegistry;
            var exp = new ExpensesViewModel(ReportDateFrom, ReportDateTo, ref db, CompanyName);
            await displayRootRegistry.ShowModalPresentation(exp);
            LoadReport.Execute("");
        }));
        public Command LaodExpensesFromXls => _loadExpensesFromXls ?? (_loadExpensesFromXls = new Command(obj =>
        {
            string path = OpenFile("Файл Excel|*.XLSX;*.XLS;*.XLSM");   // Вибираємо наш файл (метод OpenFile() описаний нижче)

            if (path == null) // Перевіряємо шлях до файлу на null
            {
                return;
            }

            if (worker != null && worker.IsBusy) { return; }
            worker = new BackgroundWorker();
            worker.DoWork += worker_DoWork;
            worker.RunWorkerAsync();
            void worker_DoWork(object sender, DoWorkEventArgs e)
            {
                ProgressStart();

                // Основні затратні задачі
                Excel.Application ExcelApp = new Excel.Application();     // Створюємо додаток Excel
                Excel.Workbook ExcelWorkBook;                             // Створюємо книгу Excel
                Excel.Worksheet ExcelWorkSheet;                           // Створюємо лист Excel            


                try
                {
                    ExcelWorkBook = ExcelApp.Workbooks.Open(path);                  // Відкриваємо файл Excel                
                    ExcelWorkSheet = ExcelWorkBook.ActiveSheet;                     // Відкриваємо активний Лист Excel                

                    for (int i = 3; i <= 1000; i++)
                    {
                        if (ExcelApp.Cells[i, 5].Value2 > 0)
                        {
                            Expenses expenses = new Expenses();

                            expenses.Date = DateTime.Parse(ExcelApp.Cells[i, 1].Value.ToString());
                            expenses.Type = ExcelApp.Cells[i, 2].Value2?.ToString();
                            expenses.Name = ExcelApp.Cells[i, 3].Value2?.ToString();
                            expenses.Description = ExcelApp.Cells[i, 4].Value2?.ToString();
                            expenses.Amounts = decimal.TryParse(ExcelApp.Cells[i, 5].Value2?.ToString(), out decimal result) ? (result) : 0m;

                            db.Expenses.Add(expenses);
                            expenses = null;
                        }
                    }
                    db.SaveChanges();


                    ExcelApp.Sheets[2].Cells[1, 20] = 1;


                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                    ExcelApp.Visible = true;
                    ExcelApp.UserControl = true;
                }
                //************************

                ProgressStop();
            }
        }));
        public Command PaymentExpensesActive => _paymentExpensesActive ?? (_paymentExpensesActive = new Command(async obj =>
        {
            if (ActiveExpenseSelect != null)
            {
                var displayRootRegistry = (Application.Current as App).displayRootRegistry;
                var pay = new DebtsPaymentViewModel();
                pay.Date = DateTime.Today;
                pay.Amount = ActiveExpenseSelect.Amounts;
                await displayRootRegistry.ShowModalPresentation(pay);

                if (pay.PressOk)
                {
                    if (pay.Amount != 0m)
                    {
                        ActiveExpenseSelect.DatePaid = pay.Date;
                        ActiveExpenseSelect.AmountPaid = decimal.Round(pay.Amount, 2);
                        ActiveExpenseSelect.NotesPaid = pay.Description;
                        ActiveExpenseSelect.Payment = true;
                        ActiveExpenseSelect.Color = "Silver";

                        db.Entry(ActiveExpenseSelect).State = EntityState.Modified;
                        db.SaveChanges();

                        ActiveExpenses.Clear();
                        PaymentExpenses.Clear();
                        ActiveExpenses = ReportActivExpenses(DateYearSelect, DateMonthSelect);
                        PaymentExpenses = ReportPaymentExpenses(DateYearSelect, DateMonthSelect);
                        LabelActiveExpenses = decimal.Round(ActiveExpenses.Select(e => e.Amounts)?.Sum() ?? 0m, 2);
                        LabelPaymentExpenses = decimal.Round(PaymentExpenses.Select(e => e.AmountPaid)?.Sum() ?? 0m, 2);
                    }
                }
            }
        }));
        public Command PaymentExpensesPayment => _paymentExpensesPayment ?? (_paymentExpensesPayment = new Command(async obj =>
        {
            if (PaymentExpenseSelect != null)
            {
                var displayRootRegistry = (Application.Current as App).displayRootRegistry;
                var pay = new DebtsPaymentViewModel();
                pay.Date = PaymentExpenseSelect.DatePaid;
                pay.Amount = PaymentExpenseSelect.AmountPaid;
                pay.Description = PaymentExpenseSelect.NotesPaid;
                await displayRootRegistry.ShowModalPresentation(pay);
                if (pay.PressOk)
                {
                    if (pay.Amount != 0m)
                    {
                        PaymentExpenseSelect.DatePaid = pay.Date;
                        PaymentExpenseSelect.AmountPaid = decimal.Round(pay.Amount, 2);
                        PaymentExpenseSelect.NotesPaid = pay.Description;
                        PaymentExpenseSelect.Payment = true;
                        PaymentExpenseSelect.Color = "Silver";

                        db.Entry(PaymentExpenseSelect).State = EntityState.Modified;
                        db.SaveChanges();

                        ActiveExpenses.Clear();
                        PaymentExpenses.Clear();
                        ActiveExpenses = ReportActivExpenses(DateYearSelect, DateMonthSelect);
                        PaymentExpenses = ReportPaymentExpenses(DateYearSelect, DateMonthSelect);
                        LabelActiveExpenses = decimal.Round(ActiveExpenses.Select(e => e.Amounts)?.Sum() ?? 0m, 2);
                        LabelPaymentExpenses = decimal.Round(PaymentExpenses.Select(e => e.AmountPaid)?.Sum() ?? 0m, 2);
                    }
                    else
                    {
                        PaymentExpenseSelect.DatePaid = DateTime.MinValue;
                        PaymentExpenseSelect.AmountPaid = 0m;
                        PaymentExpenseSelect.NotesPaid = "";
                        PaymentExpenseSelect.Payment = false;
                        PaymentExpenseSelect.Color = "Red";

                        db.Entry(PaymentExpenseSelect).State = EntityState.Modified;
                        db.SaveChanges();

                        ActiveExpenses.Clear();
                        PaymentExpenses.Clear();
                        ActiveExpenses = ReportActivExpenses(DateYearSelect, DateMonthSelect);
                        PaymentExpenses = ReportPaymentExpenses(DateYearSelect, DateMonthSelect);
                        LabelActiveExpenses = decimal.Round(ActiveExpenses.Select(e => e.Amounts)?.Sum() ?? 0m, 2);
                        LabelPaymentExpenses = decimal.Round(PaymentExpenses.Select(e => e.AmountPaid)?.Sum() ?? 0m, 2);
                    }
                }
            }
        }));
        public Command TemplateExpenses => _templateExpenses ?? (_templateExpenses = new Command(obj =>
        {
            if (worker != null && worker.IsBusy) { return; }
            worker = new BackgroundWorker();
            worker.DoWork += worker_DoWork;
            worker.RunWorkerAsync();
            void worker_DoWork(object sender, DoWorkEventArgs e)
            {
                ProgressStart();
                try
                {
                    // Основні затратні задачі
                    Excel.Application ExcelApp = new Excel.Application();
                    Excel.Workbook ExcelWorkBook;
                    ExcelWorkBook = ExcelApp.Workbooks.Open(Environment.CurrentDirectory + "\\Blanks\\TemplateExpenses");   //Вказуємо шлях до шаблону  "\\Blanks\\TemplateExpenses.xltm"

                    ExcelApp.Cells[3, 1] = DateTime.Today.ToShortDateString();

                    ExcelApp.Visible = true;           // Робим книгу видимою
                    ExcelApp.UserControl = true;       // Передаємо керування користувачу
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error message: " + Environment.NewLine +
                                           ex.Message + Environment.NewLine + Environment.NewLine +
                                           "StackTrace message: " + Environment.NewLine +
                                           ex.StackTrace, "Warning !!!");
                }

                ProgressStop();
            }
        }));
        public Command ExportReport => _exportReport ?? (_exportReport = new Command(obj =>
        {
            if (TotalReport != null)
            {
                ExportExcelTotalReport(ReportDateFrom, ReportDateTo, TotalReport);
            }
        }));
        //**********************************
        public Command DIC_TypeOfClient => _dicTypeOfClient ?? (_dicTypeOfClient = new Command(async obj =>
        {
            var displayRootRegistry = (Application.Current as App).displayRootRegistry;
            var dicClientViewModel = new DIC_ClientViewModel(ref db, EnumDictionary.TypeOfClient);
            await displayRootRegistry.ShowModalPresentation(dicClientViewModel);
        }));
        public Command DIC_HearAboutsUs => _dicHearAboutsUs ?? (_dicHearAboutsUs = new Command(async obj =>
        {
            var displayRootRegistry = (Application.Current as App).displayRootRegistry;
            var dicClientViewModel = new DIC_ClientViewModel(ref db, EnumDictionary.HearAboutsUs);
            await displayRootRegistry.ShowModalPresentation(dicClientViewModel);
        }));
        public Command DIC_Quotation => _dicQuotation ?? (_dicQuotation = new Command(async obj =>
        {
            var displayRootRegistry = (Application.Current as App).displayRootRegistry;
            var dicQuotationViewModel = new DIC_QuotationViewModel(ref db, EnumDictionary.Flooring);
            await displayRootRegistry.ShowModalPresentation(dicQuotationViewModel);
        }));
        public Command DIC_Contractors => _dicContractors ?? (_dicContractors = new Command(async obj =>
        {
            var displayRootRegistry = (Application.Current as App).displayRootRegistry;
            var dicContract = new DIC_ContractorViewModel(ref db);
            await displayRootRegistry.ShowModalPresentation(dicContract);
        }));
        public Command DIC_Supplier => _dicSupplier ?? (_dicSupplier = new Command(async obj =>
        {
            var displayRootRegistry = (Application.Current as App).displayRootRegistry;
            var supplierViewModel = new DIC_SupplierViewModel(ref db);
            await displayRootRegistry.ShowModalPresentation(supplierViewModel);
        }));
        public Command DIC_Depth => _dicDepth ?? (_dicDepth = new Command(async obj =>
        {
            var displayRootRegistry = (Application.Current as App).displayRootRegistry;
            var depthViewModel = new DIC_DepthViewModel(ref db);
            await displayRootRegistry.ShowModalPresentation(depthViewModel);
        }));

        #endregion
        public MainViewModel()
        {


            ProgressStop();

            db = new BuilderContext();
            db.Clients.Load();
            db.Quotations.Load();
            db.Invoices.Load();
            db.WorkOrders.Load();
            db.MaterialProfits.Load();
            db.LabourProfits.Load();
            db.Debts.Load();
            db.Deliveries.Load();

            LoadClientsDB();

            IsCheckedCMO = true;    // в проперті завантажуємо решту таблиць, крім Клієнтів
            IsCheckedLeveling = false;
            IsVisibleCMO = Visibility.Visible;
            IsVisibleLeveling = Visibility.Collapsed;

            ListLoaded();
            NameQuotaSelect = "ESTIMATE";

            EnableClient = false;
            EnableQuota = false;
            EnableInvoice = false;
            EnableWorkOrder = false;
            EnableMaterialProfit = false;
            EnableLabourProfit = false;
            EnableDebts = false;

            EnableReport = false;
            IsCheckedProfit = true;
            IsCheckedTotal = false;
            IsCheckedPayroll = false;
            IsCheckedDebts = false;
            IsVisibleDateSelect = Visibility.Visible;
            IsVisibleDateSelectExpenses = Visibility.Collapsed;
            IsVisibleMenuReport = Visibility.Visible;
            IsVisibleProfitReport = Visibility.Collapsed;
            IsVisibleTotalReport = Visibility.Collapsed;
            IsVisiblePayrollReport = Visibility.Collapsed;
            IsVisibleAmountReport = Visibility.Collapsed;
            IsVisibleDebtsReport = Visibility.Collapsed;
            IsVisibleExpensesReport = Visibility.Collapsed;

            ReportDateFrom = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
            ReportDateTo = new DateTime(DateTime.Today.Year, DateTime.Today.Month, DateTime.DaysInMonth(DateTime.Today.Year, DateTime.Today.Month));
        }
        private void LoadClientsDB()
        {
            Clients = db.Clients.Local.ToBindingList();
        }
        private void LoadQuotationsDB(string companyName)
        {
            Quotations = db.Quotations.Local.ToBindingList().Where(q => q.CompanyName == companyName).OrderBy(q => q.SortingQuota).ThenBy(q => q.Id);
        }
        private void LoadInvoicesDB(string companyName)
        {
            Invoices = db.Invoices.Local.ToBindingList().Where(i => i.CompanyName == companyName);
        }
        private void LoadWorkOrdersDB(string companyName)
        {
            WorkOrders = db.WorkOrders.Local.ToBindingList().Where(w => w.CompanyName == companyName);
        }
        private void LoadMaterialProfitsDB(string companyName)
        {
            MaterialProfits = db.MaterialProfits.Local.ToBindingList().Where(m => m.CompanyName == companyName);
        }
        private void LoadLabourProfitsDB(string companyName)
        {
            LabourProfits = db.LabourProfits.Local.ToBindingList().Where(l => l.CompanyName == companyName);
        }
        private void LoadDebtsDB(string companyName)
        {
            Debts = db.Debts.Local.ToBindingList().Where(d => d.CompanyName == companyName);
        }
        private void LoadDeliveriesDB(string companyName, bool isArchive)
        {
            Deliveries = db.Deliveries.Local.ToBindingList().Where(d => d.IsArchive == isArchive && d.CompanyName == companyName);
        }
        private async void StartLoadDB()
        {
            ProgressStart();
            await ClearAndLoadDB(CompanyName);
            ProgressStop();

            async Task ClearAndLoadDB(string companyName)
            {
                await Task.Run(() =>
                {
                    Quotations = null;
                    Deliveries = null;
                    WorkOrders = null;
                    Invoices = null;
                    MaterialProfits = null;
                    LabourProfits = null;
                    Debts = null;
                    LoadQuotationsDB(companyName);
                    LoadDeliveriesDB(companyName, false);
                    LoadWorkOrdersDB(companyName);
                    LoadInvoicesDB(companyName);
                    LoadMaterialProfitsDB(companyName);
                    LoadLabourProfitsDB(companyName);
                    LoadDebtsDB(companyName);
                });
            }
        }
        /// <summary>
        /// Відображає ProgressBar та робить форму напівпрозорою
        /// </summary>
        private void ProgressStart()
        {
            Progress = Visibility.Visible;
            OpacityProgress = 0.5m;
        }
        /// <summary>
        /// Приховує ProgressBar та повертає форму до нормального вигляду
        /// </summary>
        private void ProgressStop()
        {
            Progress = Visibility.Hidden;
            OpacityProgress = 1m;
        }
        /// <summary>
        /// Підраховує вартість вказаної Quota
        /// </summary>
        /// <param name="id"></param>
        private void Calculate(int? id)
        {
            var temp = db.Quotations.FirstOrDefault(q => q.Id == id);
            decimal flooring = new decimal();
            decimal accessories = new decimal();
            decimal installation = new decimal();
            decimal demolition = new decimal();
            decimal services = new decimal();
            decimal delivery = new decimal();

            if (temp != null)
            {
                var sub = db.MaterialQuotations.Where(q => q.QuotationId == QuotationSelect.Id);

                foreach (MaterialQuotation item in sub)
                {
                    switch (item.Groupe)
                    {
                        case "FLOORING":
                            {
                                flooring += item.Price;
                            }
                            break;
                        case "ACCESSORIES":
                            {
                                accessories += item.Price;
                            }
                            break;
                        case "INSTALLATION":
                            {
                                installation += item.Price;
                            }
                            break;
                        case "DEMOLITION":
                            {
                                demolition += item.Price;
                            }
                            break;
                        case "OPTIONAL SERVICES":
                            {
                                services += item.Price;
                            }
                            break;
                        case "FLOORING DELIVERY":
                            {
                                services += item.Price;
                            }
                            break;
                        default:
                            break;
                    }
                }

                temp.MaterialSubtotal = decimal.Parse((flooring + accessories).ToString());
                temp.LabourSubtotal = decimal.Parse((installation + demolition + services + delivery).ToString());

                db.Entry(temp).State = EntityState.Modified;
                db.SaveChanges();

                Quotations = null;
                LoadQuotationsDB(CompanyName);
            }
        }
        /// <summary>
        /// Перераховує баланс враховуючи зміни в Quota
        /// </summary>
        /// <param name="quota"></param>
        private void CalculatePayment(Quotation quota)
        {
            var payments = db.Payments.Where(p => p.QuotationId == quota.Id).OrderBy(p => p.Id);
            decimal quotaTotal = quota.InvoiceGrandTotal;
            foreach (var item in payments)
            {
                item.Balance = quotaTotal - item.PaymentPrincipalPaid;
                quotaTotal = item.Balance;
                db.Entry(item).State = EntityState.Modified;
                db.SaveChanges();
            }
        }
        /// <summary>
        /// Реалізує List<> та завантажує дані для відображення
        /// </summary>
        private void ListLoaded()
        {
            DeliveriesComboBox = new List<string>();
            DeliveriesComboBox = DeliveriesComboBoxGet();

            NameQuota = new List<string>();
            NameQuota.Add("ESTIMATE");
            NameQuota.Add("APPROVED QUOTE");
            NameQuota.Add("REVISED AGREEMENT");
            NameQuota.Add("INVOICE");

            DateYears = new List<int>();
            var date = Quotations?.Select(q => q.QuotaDate);
            if (date.Count() > 0)
            {
                int year = date.Min().Year;
                for (int i = 0; i <= 5; i++)
                {
                    DateYears.Add(year + i);
                }
                DateYearSelect = DateTime.Today.Year;
            }
            else
            {
                DateYears.Add(2020);
            }

            DateMonths = new List<EnumMonths>();
            for (EnumMonths month = EnumMonths.January; month <= EnumMonths.December; month++)
            {
                DateMonths.Add(month);
            }
            DateMonthSelect = DateMonths.ElementAt(DateTime.Today.Month - 1);

        }
        /// <summary>
        /// Сворює екземпляр Delivery
        /// </summary>
        /// <param name="quota"></param>
        private void AddDelivery(Quotation quota)
        {
            var material = db.MaterialQuotations.Where(m => m.QuotationId == quota.Id);
            var supplier = material?.Select(m => m.SupplierId)?.Distinct();
            if (supplier.Count() > 0)
            {
                foreach (var item in supplier)
                {
                    var suppMat = material.Where(m => m.SupplierId == item);

                    var dicSupp = db.DIC_Suppliers.Find(item);
                    if (dicSupp != null)
                    {

                        Delivery delivery = new Delivery()
                        {
                            QuotaId = quota.Id,
                            NumberQuota = quota.NumberQuota,
                            DateCreating = DateTime.Today,
                            FirstName = quota.FirstName,
                            LastName = quota.LastName,
                            PhoneNumber = quota.PhoneNumber,
                            Email = quota.Email,
                            SupplierId = dicSupp.Id,
                            SupplierName = dicSupp.Supplier,
                            CompanyName = quota.CompanyName
                        };
                        db.Deliveries.Add(delivery);
                        db.SaveChanges();
                    }
                    //OnPropertyChanged(nameof(Deliveries));
                    Deliveries = null;
                    LoadDeliveriesDB(CompanyName, false);

                    var del = db.Deliveries.OrderByDescending(d => d.Id).FirstOrDefault();

                    foreach (var Mat in suppMat)
                    {
                        if (Mat != null && del != null && dicSupp != null)
                        {
                            DeliveryMaterial deliveryMaterial = new DeliveryMaterial()
                            {
                                Groupe = Mat.Groupe,
                                Item = Mat.Item,
                                Description = Mat.Description,
                                DeliveryId = del.Id,
                                SupplierId = dicSupp.Id,
                                Supplier = dicSupp.Supplier,
                                Quantity = Mat.Quantity,
                                Rate = Mat.Rate,
                                Price = Mat.Price
                            };
                            db.DeliveryMaterials.Add(deliveryMaterial);
                            db.SaveChanges();
                        }
                    }
                }
            }
        }
        /// <summary>
        /// Видаляє екземпляр Delivery
        /// </summary>
        /// <param name="quota"></param>
        private void DelDeliveri(Quotation quota)
        {
            var delivery = db.Deliveries.Where(d => d.QuotaId == quota.Id);
            if (delivery != null)
            {
                foreach (var item in delivery)
                {
                    var material = db.DeliveryMaterials.Where(m => m.DeliveryId == item.Id);
                    db.DeliveryMaterials.RemoveRange(material);
                }
                db.Deliveries.RemoveRange(delivery);
                db.SaveChanges();
                Deliveries = null;
                LoadDeliveriesDB(CompanyName, false);
            }
        }
        /// <summary>
        /// Створює екземпляр MaterialProfit по заданому Invoice та записує в db.
        /// </summary>
        /// <param name="invoiceId"></param>
        private void AddMaterialProfit(int? invoiceId)
        {
            var invoice = db.Invoices.FirstOrDefault(i => i.Id == invoiceId);
            var quota = db.Quotations.FirstOrDefault(q => q.Id == invoice.QuotaId);
            var materialQ = db.MaterialQuotations.Where(m => m.QuotationId == quota.Id);
            MaterialProfit mp = new MaterialProfit()
            {
                InvoiceId = invoice.Id,
                InvoiceDate = invoice.DateInvoice,
                InvoiceNumber = invoice.NumberInvoice,
                FirstName = invoice.FirstName,
                LastName = invoice.LastName,
                PhoneNumber = invoice.PhoneNumber,
                Email = invoice.Email,
                MaterialSubtotal = quota.MaterialSubtotal,
                MaterialTax = quota.MaterialTax,
                MaterialTotal = decimal.Round(quota.MaterialSubtotal + quota.MaterialTax, 2),
                CostMaterialSubtotal = 0m,
                CostMaterialTax = 0m,
                CostMaterialTotal = 0m,
                ProfitBeforTax = decimal.Round(quota.MaterialSubtotal - 0m, 2),
                ProfitTax = decimal.Round(quota.MaterialTax - 0m, 2),
                ProfitInclTax = decimal.Round((quota.MaterialSubtotal + quota.MaterialTax) - 0m, 2),
                ProfitDiscount = quota.MaterialDiscountAmount,
                ProfitTotal = decimal.Round((quota.MaterialSubtotal + quota.MaterialTax) - 0m, 2) - quota.MaterialDiscountAmount,
                CompanyName = invoice.CompanyName
            };
            db.MaterialProfits.Add(mp);
            db.SaveChanges();
            MaterialProfitSelect = db.MaterialProfits.Local.ToBindingList().OrderByDescending(q => q.Id).FirstOrDefault();

            foreach (var item in materialQ)
            {
                if (item.Groupe == "FLOORING" || item.Groupe == "ACCESSORIES")
                {
                    Material mat = new Material()
                    {
                        MaterialProfitId = MaterialProfitSelect.Id,
                        Groupe = item.Groupe,
                        Item = item.Item,
                        Description = item.Description,
                        Quantity = item.Quantity,
                        Rate = item.Rate,
                        Price = item.Price,
                        SupplierId = item?.SupplierId,
                        CostQuantity = 0m,
                        CostUnitPrice = 0m,
                        CostSubtotal = 0m,
                        CostEPRate = 0m,
                        CostTax = 0m,
                        CostTotal = 0m
                    };
                    db.Materials.Add(mat);
                    db.SaveChanges();
                }
            }

            MaterialProfits = null;
            LoadMaterialProfitsDB(CompanyName);

            MaterialProfitSelect = MaterialProfits.OrderByDescending(m => m.Id).FirstOrDefault();

        }
        /// <summary>
        /// Підраховує вартість всіх матеріалів в заданому MaterialProfit
        /// </summary>
        /// <param name="select"></param>
        private void MaterialProfitCalculate(MaterialProfit select)
        {
            var material = db.Materials.Where(m => m.MaterialProfitId == select.Id);
            if (material.Count() > 0)
            {
                var price = material?.Select(m => m.Price);
                decimal priceSum = price?.Sum() ?? 0m;

                var costSubtotal = material?.Select(m => m.CostSubtotal);
                decimal costSubtotalSum = costSubtotal?.Sum() ?? 0m;

                var costTotal = material?.Select(t => t.CostTotal);
                decimal costTotalSum = costTotal?.Sum() ?? 0m;

                decimal materialSubtotal = decimal.Round(priceSum, 2);
                decimal materialTax = decimal.Round(materialSubtotal * 0.12m, 2);
                decimal materialTotal = decimal.Round(materialSubtotal + materialTax, 2);

                decimal otherSubtotal = decimal.Round(costSubtotalSum, 2);
                decimal otherTax = decimal.Round(costTotalSum - otherSubtotal, 2);
                decimal otherTotal = decimal.Round(otherSubtotal + otherTax, 2);

                decimal profitBefor = decimal.Round(materialSubtotal - otherSubtotal, 2);
                decimal profitTax = decimal.Round(materialTax - otherTax, 2);
                decimal profitInc = decimal.Round(materialTotal - otherTotal, 2);
                decimal profitTotal = decimal.Round(profitInc - (select?.ProfitDiscount ?? 0m), 2);


                select.MaterialSubtotal = materialSubtotal;
                select.MaterialTax = materialTax;
                select.MaterialTotal = materialTotal;

                select.CostMaterialSubtotal = otherSubtotal;
                select.CostMaterialTax = otherTax;
                select.CostMaterialTotal = otherTotal;

                select.ProfitBeforTax = profitBefor;
                select.ProfitTax = profitTax;
                select.ProfitInclTax = profitInc;
                select.ProfitTotal = profitTotal;

                db.Entry(select).State = EntityState.Modified;
                db.SaveChanges();

                MaterialProfits = null;
                LoadMaterialProfitsDB(CompanyName);
            }
        }
        /// <summary>
        /// Створює екземпляр LabourProfit по заданому Invoice та записує в db.
        /// </summary>
        /// <param name="invoiceId"></param>
        private void AddLabourProfit(int? invoiceId)
        {
            var invoice = db.Invoices.FirstOrDefault(i => i.Id == invoiceId);
            var quota = db.Quotations.FirstOrDefault(q => q.Id == invoice.QuotaId);
            var materialQ = db.MaterialQuotations.Where(m => m.QuotationId == quota.Id); //.Where(g=>g.Groupe == "FLOORING DELIVERY");
            var order = db.WorkOrders.FirstOrDefault(o => o.QuotaId == quota.Id);


            LabourProfit profit = new LabourProfit()
            {
                InvoiceId = invoice.Id,
                InvoiceDate = invoice.DateInvoice,
                InvoiceNumber = invoice.NumberInvoice,
                Discount = quota.LabourDiscountAmount,
                FirstName = invoice.FirstName,
                LastName = invoice.LastName,
                PhoneNumber = invoice.PhoneNumber,
                Email = invoice.Email,
                CompanyName = invoice.CompanyName
            };
            db.LabourProfits.Add(profit);
            db.SaveChanges();
            LabourProfitSelect = db.LabourProfits.Local.ToBindingList().OrderByDescending(q => q.Id).FirstOrDefault();

            if (order != null)
            {
                var install = db.WorkOrder_Installations.Where(i => i.WorkOrderId == order.Id);
                var contractor = db.WorkOrder_Contractors.Where(c => c.WorkOrderId == order.Id);
                foreach (var item in install)
                {
                    Labour lab = new Labour()
                    {
                        Groupe = item.Groupe,
                        Item = item.Item,
                        Description = item.Description,
                        Contractor = item.Contractor,
                        Quantity = item.Quantity,
                        Rate = item.Rate,
                        Price = item.Price,
                        Percent = item.Procent,
                        Payout = item.Payout,
                        Adjust = 0m,
                        Profit = decimal.Round(item.Price - item.Payout, 2),
                        LabourProfitId = LabourProfitSelect.Id
                    };
                    db.Labours.Add(lab);
                    db.SaveChanges();
                }

                foreach (var item in contractor)
                {
                    LabourContractor lc = new LabourContractor()
                    {
                        Contractor = item.Contractor,
                        Adjust = item.Adjust,
                        GST = item.GST,
                        Payout = item.Payout,
                        TAX = item.TAX,
                        Total = item.Total,
                        TotalContractor = item.TotalContractor,
                        LabourProfitId = LabourProfitSelect.Id
                    };
                    db.LabourContractors.Add(lc);
                    db.SaveChanges();
                }
            }
            else
            {
                foreach (var item in materialQ)
                {
                    if (item.Groupe == "INSTALLATION" || item.Groupe == "DEMOLITION" || item.Groupe == "OPTIONAL SERVICES" || item.Groupe == "FLOORING DELIVERY")
                    {
                        Labour lab = new Labour()
                        {
                            Groupe = item.Groupe,
                            Item = item.Item,
                            Description = item.Description,
                            Contractor = null,
                            Quantity = item.Quantity,
                            Rate = item.Rate,
                            Price = item.Price,
                            Percent = 0,
                            Payout = 0m,
                            Adjust = 0m,
                            Profit = decimal.Round(item.Price - 0m, 2),
                            LabourProfitId = LabourProfitSelect.Id
                        };
                        db.Labours.Add(lab);
                        db.SaveChanges();
                    }
                }
            }

            LabourProfits = null;
            LoadLabourProfitsDB(CompanyName);

            LabourProfitSelect = LabourProfits.OrderByDescending(l => l.Id).FirstOrDefault();
        }
        /// <summary>
        ///  Підраховує вартість всіх робіт в заданому LabourProfit
        /// </summary>
        /// <param name="select"></param>
        private void LabourProfitCalculate(LabourProfit select)
        {
            var lab = db.Labours.Where(i => i.LabourProfitId == select.Id);
            var con = db.LabourContractors.Where(c => c.LabourProfitId == select.Id);
            if (lab.Count() > 0 && con.Count() > 0)
            {
                decimal collSubtotal = decimal.Round(lab.Select(l => l.Price).Sum(), 2);
                decimal collGST = decimal.Round(collSubtotal * 0.05m, 2);
                decimal collTotal = decimal.Round(collSubtotal + collGST, 2);

                decimal paySubtotal = decimal.Round(lab.Select(l => l.Payout).Sum(), 2);
                decimal payGST = decimal.Round(con.Select(c => c.GST).Sum(), 2);
                decimal payTotal = decimal.Round(paySubtotal + payGST, 2);

                decimal storeSubtotal = decimal.Round(lab.Select(l => l.Profit).Sum(), 2);
                decimal storeGST = decimal.Round(collGST - payGST, 2);
                decimal storeTotal = decimal.Round(collTotal - payTotal, 2);

                decimal profitTotal = decimal.Round(storeTotal - select.Discount, 2);

                select.CollectedSubtotal = collSubtotal;
                select.CollectedGST = collGST;
                select.CollectedTotal = collTotal;

                select.PayoutSubtotal = paySubtotal;
                select.PayoutGST = payGST;
                select.PayoutTotal = payTotal;

                select.StoreSubtotal = storeSubtotal;
                select.StoreGST = storeGST;
                select.StoreTotal = storeTotal;

                select.ProfitTotal = profitTotal;

                db.Entry(select).State = EntityState.Modified;
                db.SaveChanges();

                LabourProfits = null;
                LoadLabourProfitsDB(CompanyName);
            }

        }
        /// <summary>
        /// Записує профіти в створений Invoice
        /// </summary>
        /// <param name="invoiceId"></param>
        private void InvoiceCalculate(int? invoiceId)
        {
            decimal material = db.MaterialProfits.FirstOrDefault(m => m.InvoiceId == invoiceId)?.ProfitTotal ?? 0m;
            decimal labour = db.LabourProfits.FirstOrDefault(l => l.InvoiceId == invoiceId)?.ProfitTotal ?? 0m;
            var invoice = db.Invoices.Find(invoiceId);
            invoice.MaterialProfit = material;
            invoice.LabourProfit = labour;
            invoice.TotalProfit = decimal.Round(material + labour, 2);
            db.Entry(invoice).State = EntityState.Modified;
            db.SaveChanges();
            Invoices = null;
            LoadInvoicesDB(CompanyName);
        }
        /// <summary>
        /// Створює екземпляри Debts по всім Contractor з вказаного Invoice та записує їх в db
        /// </summary>
        /// <param name="invoiceId"></param>
        private void AddDebtsTab(int? invoiceId)
        {
            var invoice = Invoices.FirstOrDefault(i => i.Id == invoiceId); //db.Invoices.FirstOrDefault(i => i.Id == invoiceId);
            int? idLabour = LabourProfits.FirstOrDefault(l => l.InvoiceId == invoiceId).Id; //db.LabourProfits.FirstOrDefault(l => l.InvoiceId == invoiceId).Id;
            var contractor = db.LabourContractors.Where(c => c.LabourProfitId == idLabour);
            List<Debts> debts = new List<Debts>();
            if (contractor != null)
            {
                foreach (var item in contractor)
                {
                    Debts debt = new Debts()
                    {
                        InvoiceDate = invoice.DateInvoice,
                        InvoiceId = invoice.Id,
                        InvoiceNumber = invoice.NumberInvoice,
                        FirstName = invoice.FirstName,
                        LastName = invoice.LastName,
                        Email = invoice.Email,
                        PhoneNumber = invoice.PhoneNumber,
                        NameDebts = "Payment to the contractor",
                        DescriptionDebts = item.Contractor,
                        AmountDebts = item.TotalContractor,
                        DatePayment = DateTime.MinValue,
                        CompanyName = invoice.CompanyName
                    };
                    debts.Add(debt);
                }
                db.Debts.AddRange(debts);
                db.SaveChanges();

                Debts = null;
                LoadDebtsDB(CompanyName);
            }
        }
        /// <summary>
        /// Видаляє, а потім по новій створює екземпляри Debts по всім Contractor з вказаного Invoice та записує їх в db
        /// </summary>
        /// <param name="invoiceId"></param>
        private void InsDebtsTab(int? invoiceId)
        {
            var invoice = Invoices.FirstOrDefault(i => i.Id == invoiceId); //db.Invoices.FirstOrDefault(i => i.Id == invoiceId);
            int? idLabour = LabourProfits.FirstOrDefault(l => l.InvoiceId == invoiceId).Id; //db.LabourProfits.FirstOrDefault(l => l.InvoiceId == invoiceId).Id;
            var contractor = db.LabourContractors.Where(c => c.LabourProfitId == idLabour);

            var labourDebts = Debts.Where(d => d.InvoiceId == invoiceId && d.ReadOnly == true); //db.Debts.Where(d => d.InvoiceId == invoiceId && d.ReadOnly == true);
            db.Debts.RemoveRange(labourDebts);
            db.SaveChanges();

            List<Debts> debts = new List<Debts>();
            if (contractor != null)
            {
                foreach (var item in contractor)
                {
                    Debts debt = new Debts()
                    {
                        InvoiceDate = invoice.DateInvoice,
                        InvoiceId = invoice.Id,
                        InvoiceNumber = invoice.NumberInvoice,
                        FirstName = invoice.FirstName,
                        LastName = invoice.LastName,
                        Email = invoice.Email,
                        PhoneNumber = invoice.PhoneNumber,
                        NameDebts = "Payment to the contractor",
                        DescriptionDebts = item.Contractor,
                        AmountDebts = item.TotalContractor,
                        DatePayment = DateTime.MinValue,
                        CompanyName = invoice.CompanyName
                    };
                    debts.Add(debt);
                }
                db.Debts.AddRange(debts);
                db.SaveChanges();
                Debts = null;
                LoadDebtsDB(CompanyName);
            }
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
        /// Відкриває вибрані файли по заданій масці (масив файлів)
        /// </summary>
        /// <param name="filter"></param>
        /// <returns></returns>
        private string[] OpenFileArray(string filter)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = true;
            dialog.Filter = filter;
            if (dialog.ShowDialog() == true)
            {
                return dialog.FileNames;
            }
            return null;
        }
        /// <summary>
        /// Метод закриває программу
        /// </summary>        
        private void ExitApplication()
        {
            db.Dispose();
            Application app = Application.Current;
            app.Shutdown();
        }
        /// <summary>
        /// Використовується для звіту по заданому проміжку дат. В полі InvoiceNumber повертає кількість рахунків за даний період.
        /// </summary>
        /// <param name="dateFrom"></param>
        /// <param name="dateTo"></param>
        /// <returns></returns>
        private MaterialProfit ReportMaterial(DateTime dateFrom, DateTime dateTo)
        {
            var filterMaterial = MaterialProfits.Where(m => m.InvoiceDate >= dateFrom && m.InvoiceDate <= dateTo && m.Companion == true); //db.MaterialProfits.Where(m => m.InvoiceDate >= dateFrom && m.InvoiceDate <= dateTo && m.Companion == true);
            if (filterMaterial.Count() > 0)
            {
                MaterialProfit report = new MaterialProfit();
                report.InvoiceDate = DateTime.MinValue;
                report.InvoiceNumber = filterMaterial.Count().ToString();

                report.MaterialSubtotal = decimal.Round(filterMaterial.Select(m => m.MaterialSubtotal).Sum(), 2);
                report.MaterialTax = decimal.Round(filterMaterial.Select(m => m.MaterialTax).Sum(), 2);
                report.MaterialTotal = decimal.Round(filterMaterial.Select(m => m.MaterialTotal).Sum(), 2);

                report.CostMaterialSubtotal = decimal.Round(filterMaterial.Select(m => m.CostMaterialSubtotal).Sum(), 2);
                report.CostMaterialTax = decimal.Round(filterMaterial.Select(m => m.CostMaterialTax).Sum(), 2);
                report.CostMaterialTotal = decimal.Round(filterMaterial.Select(m => m.CostMaterialTotal).Sum(), 2);

                report.ProfitBeforTax = decimal.Round(filterMaterial.Select(m => m.ProfitBeforTax).Sum(), 2);
                report.ProfitTax = decimal.Round(filterMaterial.Select(m => m.ProfitTax).Sum(), 2);
                report.ProfitInclTax = decimal.Round(filterMaterial.Select(m => m.ProfitInclTax).Sum(), 2);
                report.ProfitDiscount = decimal.Round(filterMaterial.Select(m => m.ProfitDiscount).Sum(), 2);
                report.ProfitTotal = decimal.Round(filterMaterial.Select(m => m.ProfitTotal).Sum(), 2);
                return report;
            }
            else
            {
                return null;
            }

        }
        /// <summary>
        /// Використовується для звіту по заданому проміжку дат. В полі InvoiceNumber повертає кількість рахунків за даний період.
        /// </summary>
        /// <param name="dateFrom"></param>
        /// <param name="dateTo"></param>
        /// <returns></returns>
        private LabourProfit ReportLabour(DateTime dateFrom, DateTime dateTo)
        {
            var filterLabour = LabourProfits.Where(l => l.InvoiceDate >= dateFrom && l.InvoiceDate <= dateTo && l.Companion == true); //db.LabourProfits.Where(l => l.InvoiceDate >= dateFrom && l.InvoiceDate <= dateTo && l.Companion == true);
            if (filterLabour.Count() > 0)
            {
                LabourProfit report = new LabourProfit();
                report.InvoiceDate = DateTime.MinValue;
                report.InvoiceNumber = filterLabour.Count().ToString();

                report.CollectedSubtotal = decimal.Round(filterLabour.Select(l => l.CollectedSubtotal).Sum(), 2);
                report.CollectedGST = decimal.Round(filterLabour.Select(l => l.CollectedGST).Sum(), 2);
                report.CollectedTotal = decimal.Round(filterLabour.Select(l => l.CollectedTotal).Sum(), 2);

                report.PayoutSubtotal = decimal.Round(filterLabour.Select(l => l.PayoutSubtotal).Sum(), 2);
                report.PayoutGST = decimal.Round(filterLabour.Select(l => l.PayoutGST).Sum(), 2);
                report.PayoutTotal = decimal.Round(filterLabour.Select(l => l.PayoutTotal).Sum(), 2);

                report.StoreSubtotal = decimal.Round(filterLabour.Select(l => l.StoreSubtotal).Sum(), 2);
                report.StoreGST = decimal.Round(filterLabour.Select(l => l.StoreGST).Sum(), 2);
                report.StoreTotal = decimal.Round(filterLabour.Select(l => l.StoreTotal).Sum(), 2);
                report.Discount = decimal.Round(filterLabour.Select(l => l.Discount).Sum(), 2);
                report.ProfitTotal = decimal.Round(filterLabour.Select(l => l.ProfitTotal).Sum(), 2);
                return report;
            }
            else
            {
                return null;
            }
        }
        /// <summary>
        /// Використовується для звіту по заданому проміжку дат.
        /// </summary>
        /// <param name="dateFrom"></param>
        /// <param name="dateTo"></param>
        /// <returns></returns>
        private List<Report> ReportTotalGrid(DateTime dateFrom, DateTime dateTo)
        {
            var invoice = Invoices.Where(i => i.DateInvoice >= dateFrom && i.DateInvoice <= dateTo); //db.Invoices.Where(i => i.DateInvoice >= dateFrom && i.DateInvoice <= dateTo);
            if (invoice.Count() > 0)
            {
                var material = MaterialProfits.Where(i => i.InvoiceDate >= dateFrom && i.InvoiceDate <= dateTo && i.Companion == true); //db.MaterialProfits.Where(i => i.InvoiceDate >= dateFrom && i.InvoiceDate <= dateTo && i.Companion == true);
                var labour = LabourProfits.Where(i => i.InvoiceDate >= dateFrom && i.InvoiceDate <= dateTo && i.Companion == true); //db.LabourProfits.Where(i => i.InvoiceDate >= dateFrom && i.InvoiceDate <= dateTo && i.Companion == true);

                List<Report> reports = new List<Report>();
                foreach (var item in invoice)
                {
                    var temp = db.Payments.Where(p => p.QuotationId == item.QuotaId);
                    decimal fee;
                    if (temp.Count() > 0)
                    {
                        fee = temp.Select(f => f.ProcessingFee).Sum();
                    }
                    else
                    {
                        fee = 0m;
                    }
                    Report report = new Report()
                    {
                        InvoiceNumber = item.NumberInvoice,
                        InvoiceDate = item.DateInvoice,
                        PrimaryName = item.FullName,
                        MaterialSubtotal = material.FirstOrDefault(m => m.InvoiceId == item.Id)?.MaterialSubtotal ?? 0m,
                        MaterialTax = material.FirstOrDefault(m => m.InvoiceId == item.Id)?.MaterialTax ?? 0m,
                        MaterialGrandTotal = material.FirstOrDefault(m => m.InvoiceId == item.Id)?.MaterialTotal ?? 0m,
                        LabourSubtotal = labour.FirstOrDefault(l => l.InvoiceId == item.Id)?.CollectedSubtotal ?? 0m,
                        LabourGST = labour.FirstOrDefault(l => l.InvoiceId == item.Id)?.CollectedGST ?? 0m,
                        LabourGrandTotal = labour.FirstOrDefault(l => l.InvoiceId == item.Id)?.CollectedTotal ?? 0m,
                        MaterialAndLabourGrandTotal = decimal.Round((material.FirstOrDefault(m => m.InvoiceId == item.Id)?.MaterialTotal ?? 0m)
                                                    + (labour.FirstOrDefault(l => l.InvoiceId == item.Id)?.CollectedTotal ?? 0m), 2),
                        MaterialCostSubtotal = material.FirstOrDefault(m => m.InvoiceId == item.Id)?.CostMaterialSubtotal ?? 0m,
                        MaterialCostTax = material.FirstOrDefault(m => m.InvoiceId == item.Id)?.CostMaterialTax ?? 0m,
                        MaterialDiscountDeductions = material.FirstOrDefault(m => m.InvoiceId == item.Id)?.ProfitDiscount ?? 0m,
                        MaterialCostTotal = material.FirstOrDefault(m => m.InvoiceId == item.Id)?.CostMaterialTotal ?? 0m,
                        LabourPayoutTotalBeforeGST = labour.FirstOrDefault(l => l.InvoiceId == item.Id)?.PayoutSubtotal ?? 0m,
                        LabourPayoutGST = labour.FirstOrDefault(l => l.InvoiceId == item.Id)?.PayoutGST ?? 0m,
                        LabourDiscountDeductions = labour.FirstOrDefault(l => l.InvoiceId == item.Id)?.Discount ?? 0m,
                        LabourPayoutTotalAfterGST = labour.FirstOrDefault(l => l.InvoiceId == item.Id)?.PayoutTotal ?? 0m,
                        TotalMaterialProfits = material.FirstOrDefault(m => m.InvoiceId == item.Id)?.ProfitTotal ?? 0m,
                        TotalLabourProfits = labour.FirstOrDefault(l => l.InvoiceId == item.Id)?.ProfitTotal ?? 0m,
                        ProcessingFeeCollected = fee,
                        TotalProfit = decimal.Round((material.FirstOrDefault(m => m.InvoiceId == item.Id)?.ProfitTotal ?? 0m)
                                      + (labour.FirstOrDefault(l => l.InvoiceId == item.Id)?.ProfitTotal ?? 0m), 2)
                    };
                    reports.Add(report);
                }

                return reports;
            }
            else
            {
                return null;
            }
        }
        /// <summary>
        /// Використовується для показу в полях сумми по кожній позиції звіту. По заданому проміжку дат вибираються затрати.
        /// </summary>
        /// <param name="dateFrom"></param>
        /// <param name="dateTo"></param>
        /// <param name="list"></param>
        /// <returns></returns>
        private ReportTotal ReportTotalLabel(DateTime dateFrom, DateTime dateTo, List<Report> list)
        {
            if (list != null)
            {
                var expens = db.Expenses.Where(i => i.Date >= dateFrom && i.Date <= dateTo);
                decimal expensSum = (expens.Count() > 0) ? (decimal.Round(expens.Select(e => e.Amounts).Sum(), 2)) : 0m;
                ReportTotal report = new ReportTotal()
                {
                    TotalSale = decimal.Round(list.Select(l => l.MaterialAndLabourGrandTotal).Sum(), 2),
                    TotalMaterialAndLabourCost = decimal.Round(list.Select(l => l.MaterialCostTotal).Sum() + list.Select(l => l.LabourPayoutTotalAfterGST).Sum(), 2),
                    TotalExpenses = expensSum,
                    TotalProcessFees = decimal.Round(list.Select(l => l.ProcessingFeeCollected).Sum(), 2),
                    NetProfit = decimal.Round(list.Select(l => l.TotalProfit).Sum() - expensSum, 2)
                };

                return report;
            }
            else
            {
                return null;
            }
        }
        /// <summary>
        /// Експорт звіту в шаблон Excel (Report.xltm)
        /// </summary>
        /// <param name="dateFrom"></param>
        /// <param name="dateTo"></param>
        /// <param name="reports"></param>
        private void ExportExcelTotalReport(DateTime dateFrom, DateTime dateTo, List<Report> reports)
        {
            Excel.Application ExcelApp = new Excel.Application();
            Excel.Workbook ExcelWorkBook;
            ExcelWorkBook = ExcelApp.Workbooks.Open(Environment.CurrentDirectory + "\\Blanks\\Report.xltm");   //Вказуємо шлях до шаблону

            var expens = db.Expenses.Where(i => i.Date >= dateFrom && i.Date <= dateTo);

            int counter;
            if (reports.Count > 0)
            {
                counter = 3;
                foreach (var item in reports)
                {
                    ExcelApp.Cells[counter, 1] = item.InvoiceNumber;
                    ExcelApp.Cells[counter, 2] = item.InvoiceDate;
                    ExcelApp.Cells[counter, 3] = item.PrimaryName;
                    ExcelApp.Cells[counter, 4] = item.MaterialSubtotal;
                    ExcelApp.Cells[counter, 5] = item.MaterialTax;
                    ExcelApp.Cells[counter, 6] = item.MaterialGrandTotal;
                    ExcelApp.Cells[counter, 7] = item.LabourSubtotal;
                    ExcelApp.Cells[counter, 8] = item.LabourGST;
                    ExcelApp.Cells[counter, 9] = item.LabourGrandTotal;
                    ExcelApp.Cells[counter, 10] = item.MaterialAndLabourGrandTotal;
                    ExcelApp.Cells[counter, 11] = item.MaterialCostSubtotal;
                    ExcelApp.Cells[counter, 12] = item.MaterialCostTax;
                    ExcelApp.Cells[counter, 13] = item.MaterialDiscountDeductions;
                    ExcelApp.Cells[counter, 14] = item.MaterialCostTotal;
                    ExcelApp.Cells[counter, 15] = item.LabourPayoutTotalBeforeGST;
                    ExcelApp.Cells[counter, 16] = item.LabourPayoutGST;
                    ExcelApp.Cells[counter, 17] = item.LabourDiscountDeductions;
                    ExcelApp.Cells[counter, 18] = item.LabourPayoutTotalAfterGST;
                    ExcelApp.Cells[counter, 19] = item.TotalMaterialProfits;
                    ExcelApp.Cells[counter, 20] = item.TotalLabourProfits;
                    ExcelApp.Cells[counter, 21] = item.ProcessingFeeCollected;
                    counter++;
                }
            }

            if (expens.Count() > 0)
            {
                List<Expenses> expensGruop = new List<Expenses>();

                var typeExpense = expens.Select(e => e.Type).Distinct();
                foreach (var item in typeExpense)
                {
                    var groupe = expens.Where(e => e.Type == item);
                    foreach (var query in groupe)
                    {
                        expensGruop.Add(query);
                    }
                }

                ExcelApp.Cells[3, 28] = decimal.Round(expens.Select(e => e.Amounts).Sum(), 2);

                counter = 4;
                foreach (var item in expensGruop)
                {
                    ExcelApp.Cells[counter, 24] = item.Date;
                    ExcelApp.Cells[counter, 25] = item.Type;
                    ExcelApp.Cells[counter, 26] = item.Name;
                    ExcelApp.Cells[counter, 27] = item.Description;
                    ExcelApp.Cells[counter, 28] = item.Amounts;
                    counter++;
                }
            }
            ExcelApp.Visible = true;           // Робим книгу видимою
            ExcelApp.UserControl = true;       // Передаємо керування користувачу  
        }
        /// <summary>
        /// Використовується для звіту по заданому проміжку дат. Вибирає всіх підрідників з привязкою до LabourProfit
        /// </summary>
        /// <param name="dateFrom"></param>
        /// <param name="dateTo"></param>
        /// <returns></returns>
        private List<ReportPayrollToWork> ReportPayrollWork(DateTime dateFrom, DateTime dateTo)
        {
            List<ReportPayrollToWork> reports = new List<ReportPayrollToWork>();
            //var workOrder = db.WorkOrders.Where(w => w.DateWork >= dateFrom && w.DateWork <= dateTo);
            var workOrder = LabourProfits.Where(w => w.InvoiceDate >= dateFrom && w.InvoiceDate <= dateTo && w.Companion == true); //db.LabourProfits.Where(w => w.InvoiceDate >= dateFrom && w.InvoiceDate <= dateTo && w.Companion == true);
            if (workOrder.Count() > 0)
            {
                foreach (var item in workOrder)
                {
                    foreach (var contractor in item.LabourContractors.OrderBy(c => c.Contractor))
                    {
                        ReportPayrollToWork payrollToWork = new ReportPayrollToWork()
                        {
                            NumberInvoice = item.InvoiceNumber,
                            DateInvoice = item.InvoiceDate,
                            Contractor = contractor.Contractor,
                            Adjust = contractor.Adjust,
                            Payout = contractor.Payout,
                            GST = contractor.GST,
                            TAX = contractor.TAX,
                            Total = contractor.Total,
                            TotalContractor = contractor.TotalContractor
                        };
                        reports.Add(payrollToWork);
                    }
                }
                return reports;
            }
            else
            {
                return null;
            }
        }
        /// <summary>
        /// Використовується для звіту з колекції ReportPayrollToWork. Сумує всі поля з однаковими підрядниками.
        /// </summary>
        /// <param name="payrollToWorks"></param>
        /// <returns></returns>
        private List<ReportPayroll> ReportPayroll(List<ReportPayrollToWork> payrollToWorks)
        {
            List<ReportPayroll> reports = new List<ReportPayroll>();

            if (payrollToWorks != null)
            {
                var name = payrollToWorks.Select(n => n.Contractor).Distinct();
                foreach (var item in name)
                {
                    ReportPayroll contractor = new ReportPayroll()
                    {
                        Contractor = item,
                        Adjust = decimal.Round(payrollToWorks.Where(p => p.Contractor == item).Select(p => p.Adjust).Sum(), 2),
                        GST = decimal.Round(payrollToWorks.Where(p => p.Contractor == item).Select(p => p.GST).Sum(), 2),
                        Payout = decimal.Round(payrollToWorks.Where(p => p.Contractor == item).Select(p => p.Payout).Sum(), 2),
                        Total = decimal.Round(payrollToWorks.Where(p => p.Contractor == item).Select(p => p.Total).Sum(), 2),
                        TotalContractor = decimal.Round(payrollToWorks.Where(p => p.Contractor == item).Select(p => p.TotalContractor).Sum(), 2),
                    };
                    reports.Add(contractor);
                }
                return reports.OrderBy(r => r.Contractor).ToList();
            }
            else
            {
                return null;
            }
        }
        /// <summary>
        ///  Використовується для звіту по заданому проміжку дат. Вибирає всі платежі з привязкою до Quota
        /// </summary>
        /// <param name="dateFrom"></param>
        /// <param name="dateTo"></param>
        /// <returns></returns>
        private List<ReportAmount> ReportAmountGet(DateTime dateFrom, DateTime dateTo)
        {
            var quotations = Quotations.Where(q => q.Id > 0); //db.Quotations.Where(q => q.Id > 0);

            List<ReportAmount> reports = new List<ReportAmount>();
            if (quotations.Count() > 0)
            {
                foreach (var quota in quotations)
                {
                    var payment = db.Payments.Where(p => p.QuotationId == quota.Id && p.PaymentDatePaid >= dateFrom && p.PaymentDatePaid <= dateTo);
                    foreach (var item in payment)
                    {
                        ReportAmount amount = new ReportAmount()
                        {
                            QuotaDate = quota.QuotaDate,
                            NumberQuota = quota.NumberQuota,
                            FullNameQuota = quota.FullName,
                            PaymentDatePaid = item.PaymentDatePaid,
                            PaymentAmountPaid = item.PaymentAmountPaid,
                            PaymentMethod = item.PaymentMethod,
                            PaymentPrincipalPaid = item.PaymentPrincipalPaid,
                            ProcessingFee = item.ProcessingFee,
                            Balance = item.Balance
                        };
                        reports.Add(amount);
                    }
                }
            }
            return reports?.OrderBy(r => r.PaymentDatePaid).ToList();
        }
        /// <summary>
        /// Використовується для звіту по заданому проміжку дат. Вибирає всі несплачені платежі з Debts.
        /// </summary>
        /// <param name="dateFrom"></param>
        /// <param name="dateTo"></param>
        /// <returns></returns>
        private List<Debts> ReportAmountDebts(DateTime dateFrom, DateTime dateTo)
        {
            return Debts.Where(d => d.InvoiceDate >= dateFrom && d.InvoiceDate <= dateTo && d.Payment == false).ToList(); //db.Debts.Where(d => d.InvoiceDate >= dateFrom && d.InvoiceDate <= dateTo && d.Payment == false).ToList();
        }
        /// <summary>
        /// Використовується для звіту по заданому проміжку дат. Вибирає всі оплачені платежі з Debts.
        /// </summary>
        /// <param name="dateFrom"></param>
        /// <param name="dateTo"></param>
        /// <returns></returns>
        private List<Debts> ReportPaymentDebts(DateTime dateFrom, DateTime dateTo)
        {
            return Debts.Where(d => d.InvoiceDate >= dateFrom && d.InvoiceDate <= dateTo && d.Payment == true).ToList(); //db.Debts.Where(d => d.InvoiceDate >= dateFrom && d.InvoiceDate <= dateTo && d.Payment == true).ToList();
        }
        /// <summary>
        /// Використовується для звіту по заданому Year and Month. Вибирає всі неоплачені записи з Expenses.
        /// </summary>
        /// <param name="year"></param>
        /// <param name="month"></param>
        /// <returns></returns>
        private List<Expenses> ReportActivExpenses(int year, EnumMonths month)
        {
            return db.Expenses.Where(e => e.CompanyName == CompanyName && e.Date.Year == year && e.Date.Month == (int)month && e.Payment == false).OrderBy(e => e.Type).ThenBy(e => e.Name).ToList();
        }
        /// <summary>
        /// Використовується для звіту по заданому Year and Month. Вибирає всі оплачені записи з Expenses.
        /// </summary>
        /// <param name="year"></param>
        /// <param name="month"></param>
        /// <returns></returns>
        private List<Expenses> ReportPaymentExpenses(int year, EnumMonths month)
        {
            return db.Expenses.Where(e => e.CompanyName == CompanyName && e.Date.Year == year && e.Date.Month == (int)month && e.Payment == true).OrderBy(e => e.Type).ThenBy(e => e.Name).ToList();
        }
        /// <summary>
        /// Завантажує ComboBox унікальними даними
        /// </summary>
        /// <returns></returns>
        private List<string> DeliveriesComboBoxGet()
        {
            List<string> comboBox = new List<string>();
            var name = db.Deliveries.Where(d => d.IsArchive == false).ToList();

            if (name.Count > 0)
            {
                foreach (var item in name)
                {
                    comboBox.Add(item.NameComboBox);
                }
                comboBox = comboBox.OrderBy(x => name).Distinct().ToList();
            }
            else
            {
                comboBox?.Clear();
            }

            return comboBox;
        }
        /// <summary>
        /// Завантажує ComboBox унікальними даними з Архіву
        /// </summary>
        /// <returns></returns>
        private List<string> DeliveriesComboBoxArchiveGet()
        {
            List<string> comboBox = new List<string>();
            var name = db.Deliveries.Where(d => d.IsArchive == true).ToList();

            if (name.Count > 0)
            {
                foreach (var item in name)
                {
                    comboBox.Add(item.NameComboBox);
                }
                comboBox = comboBox.OrderBy(x => name).Distinct().ToList();
            }
            else
            {
                comboBox?.Clear();
            }

            return comboBox;
        }
        /// <summary>
        /// Формує файл Excel з шаблону на 2 листи
        /// </summary>
        /// <param name="contractor"></param>
        /// <param name="path"></param>
        private void WorkOrderPrintMinToExcel(WorkOrder_Contractor contractor, string path)
        {

            Excel.Application ExcelApp = new Excel.Application();
            Excel.Workbook ExcelWorkBook;
            ExcelWorkBook = ExcelApp.Workbooks.Open(Environment.CurrentDirectory + path);   //Вказуємо шлях до шаблону


            var quota = Quotations.FirstOrDefault(q => q.Id == WorkOrderSelect.QuotaId); //db.Quotations.FirstOrDefault(q => q.Id == WorkOrderSelect.QuotaId);
            var client = Clients.FirstOrDefault(c => c.Id == quota.ClientId); //db.Clients.FirstOrDefault(c => c.Id == quota.ClientId);
            var work = db.WorkOrder_Works.Where(w => w.WorkOrderId == WorkOrderSelect.Id && w.Contractor == contractor.Contractor);
            var accessories = db.WorkOrder_Accessories.Where(a => a.WorkOrderId == WorkOrderSelect.Id && a.Contractor == contractor.Contractor);
            var inst = db.WorkOrder_Installations.Where(ins => ins.WorkOrderId == WorkOrderSelect.Id && ins.Contractor == contractor.Contractor);
            var cont = db.WorkOrder_Contractors.Where(c => c.WorkOrderId == WorkOrderSelect.Id && c.Contractor == contractor.Contractor).OrderBy(c => c.Contractor);

            ExcelApp.Cells[3, 2] = quota.JobDescription;
            ExcelApp.Cells[4, 2] = quota.NumberQuota;
            ExcelApp.Cells[5, 2] = client.CompanyName;
            ExcelApp.Cells[6, 2] = client.PrimaryFirstName + " " + client.PrimaryLastName;
            ExcelApp.Cells[7, 2] = client.PrimaryPhoneNumber;
            ExcelApp.Cells[8, 2] = client.SecondaryFirstName + " " + client.SecondaryLastName;
            ExcelApp.Cells[9, 2] = client.SecondaryPhoneNumber;
            ExcelApp.Cells[4, 6] = client.NumberClient;
            ExcelApp.Cells[5, 6] = client.AddressSiteStreet + ", " + client.AddressSiteCity + ", " + client.AddressSiteProvince + ", " + client.AddressSitePostalCode + ", " + client.AddressSiteCountry;
            ExcelApp.Cells[7, 6] = WorkOrderSelect.Parking;
            ExcelApp.Cells[8, 6] = WorkOrderSelect.DateServices;
            ExcelApp.Cells[9, 6] = WorkOrderSelect.DateCompletion;
            ExcelApp.Cells[11, 2] = WorkOrderSelect.Notes; //quota.JobNote;

            //ExcelApp.Cells[17, 3] = WorkOrderSelect.Trim;
            //ExcelApp.Cells[17, 6] = WorkOrderSelect.Colour;
            //ExcelApp.Cells[17, 8] = WorkOrderSelect.Notes;
            //ExcelApp.Cells[18, 3] = WorkOrderSelect.Baseboard;
            //ExcelApp.Cells[18, 7] = WorkOrderSelect.ReplacingYesNo;
            //ExcelApp.Cells[18, 8] = WorkOrderSelect.ReplacingQuantity;

            int i = 21;
            foreach (var item in work)
            {
                ExcelApp.Cells[i, 1] = item.Area;
                ExcelApp.Cells[i, 2] = item.Room;
                ExcelApp.Cells[i, 3] = item.Existing;
                ExcelApp.Cells[i, 5] = item.NewFloor;
                ExcelApp.Cells[i, 8] = item.Furniture;
                i++;
                ExcelApp.Cells[i, 1] = "Caution / Notes: ";
                ExcelApp.Cells[i, 2] = item.Misc;
                i++;
            }

            i = 31;
            foreach (var item in accessories)
            {
                ExcelApp.Cells[i, 1] = item.Area;
                ExcelApp.Cells[i, 2] = item.Room;
                ExcelApp.Cells[i, 3] = item.OldAccessories;
                ExcelApp.Cells[i, 5] = item.NewAccessories;
                i++;
                ExcelApp.Cells[i, 1] = "Notes: ";
                ExcelApp.Cells[i, 2] = item.Notes;
                i++;
            }

            i = 48;
            foreach (var item in inst)
            {
                if (item.Groupe == "DEMOLITION")
                {
                    ExcelApp.Cells[i, 1] = item.Item;
                    ExcelApp.Cells[i, 2] = item.Description;
                    ExcelApp.Cells[i, 5] = item.Contractor;
                    ExcelApp.Cells[i, 6] = item.Quantity;
                    ExcelApp.Cells[i, 7] = item.Procent;
                    ExcelApp.Cells[i, 8] = item.Payout;
                    i++;
                }
            }

            i = 54;
            foreach (var item in inst)
            {
                if (item.Groupe == "INSTALLATION")
                {
                    ExcelApp.Cells[i, 1] = item.Item;
                    ExcelApp.Cells[i, 2] = item.Description;
                    ExcelApp.Cells[i, 5] = item.Contractor;
                    ExcelApp.Cells[i, 6] = item.Quantity;
                    ExcelApp.Cells[i, 7] = item.Procent;
                    ExcelApp.Cells[i, 8] = item.Payout;
                    i++;
                }
            }

            i = 63;
            foreach (var item in inst)
            {
                if (item.Groupe == "OPTIONAL SERVICES")
                {
                    ExcelApp.Cells[i, 1] = item.Item;
                    ExcelApp.Cells[i, 2] = item.Description;
                    ExcelApp.Cells[i, 5] = item.Contractor;
                    ExcelApp.Cells[i, 6] = item.Quantity;
                    ExcelApp.Cells[i, 7] = item.Procent;
                    ExcelApp.Cells[i, 8] = item.Payout;
                    i++;
                }
            }

            i = 72;
            foreach (var item in cont)
            {
                ExcelApp.Cells[i, 1] = item.Contractor;
                ExcelApp.Cells[i, 2] = item.Payout;
                ExcelApp.Cells[i, 3] = item.Adjust;
                ExcelApp.Cells[i, 4] = item.Total;
                ExcelApp.Cells[i, 5] = item.TAX;
                ExcelApp.Cells[i, 6] = item.GST;
                ExcelApp.Cells[i, 7] = item.TotalContractor;
                i++;
            }

            ExcelApp.Calculate();
            ExcelApp.Cells[3, 12] = cont.FirstOrDefault().Contractor;
            ExcelApp.Cells[1, 12] = "1";   // Записуємо дані в .pdf                       
            ExcelApp.Calculate();

        }
        /// <summary>
        /// Формує файл Excel з шаблону на 3 листи
        /// </summary>
        /// <param name="contractor"></param>
        /// <param name="path"></param>
        private void WorkOrderPrintMaxToExcel(WorkOrder_Contractor contractor, string path)
        {

            Excel.Application ExcelApp = new Excel.Application();
            Excel.Workbook ExcelWorkBook;
            ExcelWorkBook = ExcelApp.Workbooks.Open(Environment.CurrentDirectory + path);   //Вказуємо шлях до шаблону


            var quota = Quotations.FirstOrDefault(q => q.Id == WorkOrderSelect.QuotaId); //db.Quotations.FirstOrDefault(q => q.Id == WorkOrderSelect.QuotaId);
            var client = Clients.FirstOrDefault(c => c.Id == quota.ClientId); //db.Clients.FirstOrDefault(c => c.Id == quota.ClientId);
            var work = db.WorkOrder_Works.Where(w => w.WorkOrderId == WorkOrderSelect.Id && w.Contractor == contractor.Contractor);
            var accessories = db.WorkOrder_Accessories.Where(a => a.WorkOrderId == WorkOrderSelect.Id && a.Contractor == contractor.Contractor);
            var inst = db.WorkOrder_Installations.Where(ins => ins.WorkOrderId == WorkOrderSelect.Id && ins.Contractor == contractor.Contractor);
            var cont = db.WorkOrder_Contractors.Where(c => c.WorkOrderId == WorkOrderSelect.Id && c.Contractor == contractor.Contractor).OrderBy(c => c.Contractor);

            ExcelApp.Cells[3, 2] = quota.JobDescription;
            ExcelApp.Cells[4, 2] = quota.NumberQuota;
            ExcelApp.Cells[5, 2] = client.CompanyName;
            ExcelApp.Cells[6, 2] = client.PrimaryFirstName + " " + client.PrimaryLastName;
            ExcelApp.Cells[7, 2] = client.PrimaryPhoneNumber;
            ExcelApp.Cells[8, 2] = client.SecondaryFirstName + " " + client.SecondaryLastName;
            ExcelApp.Cells[9, 2] = client.SecondaryPhoneNumber;
            ExcelApp.Cells[4, 6] = client.NumberClient;
            ExcelApp.Cells[5, 6] = client.AddressSiteStreet + ", " + client.AddressSiteCity + ", " + client.AddressSiteProvince + ", " + client.AddressSitePostalCode + ", " + client.AddressSiteCountry;
            ExcelApp.Cells[7, 6] = WorkOrderSelect.Parking;
            ExcelApp.Cells[8, 6] = WorkOrderSelect.DateServices;
            ExcelApp.Cells[9, 6] = WorkOrderSelect.DateCompletion;
            ExcelApp.Cells[11, 2] = WorkOrderSelect.Notes; //quota.JobNote;

            //ExcelApp.Cells[17, 3] = WorkOrderSelect.Trim;
            //ExcelApp.Cells[17, 6] = WorkOrderSelect.Colour;
            //ExcelApp.Cells[17, 8] = WorkOrderSelect.Notes;
            //ExcelApp.Cells[18, 3] = WorkOrderSelect.Baseboard;
            //ExcelApp.Cells[18, 7] = WorkOrderSelect.ReplacingYesNo;
            //ExcelApp.Cells[18, 8] = WorkOrderSelect.ReplacingQuantity;

            int i = 21;
            foreach (var item in work)
            {
                ExcelApp.Cells[i, 1] = item.Area;
                ExcelApp.Cells[i, 2] = item.Room;
                ExcelApp.Cells[i, 3] = item.Existing;
                ExcelApp.Cells[i, 5] = item.NewFloor;
                ExcelApp.Cells[i, 8] = item.Furniture;
                i++;
                ExcelApp.Cells[i, 1] = "Caution / Notes: ";
                ExcelApp.Cells[i, 2] = item.Misc;
                i++;
            }

            i = 41;
            foreach (var item in accessories)
            {
                ExcelApp.Cells[i, 1] = item.Area;
                ExcelApp.Cells[i, 2] = item.Room;
                ExcelApp.Cells[i, 3] = item.OldAccessories;
                ExcelApp.Cells[i, 5] = item.NewAccessories;
                i++;
                ExcelApp.Cells[i, 1] = "Notes: ";
                ExcelApp.Cells[i, 2] = item.Notes;
                i++;
            }

            i = 77;
            foreach (var item in inst)
            {
                if (item.Groupe == "DEMOLITION")
                {
                    ExcelApp.Cells[i, 1] = item.Item;
                    ExcelApp.Cells[i, 2] = item.Description;
                    ExcelApp.Cells[i, 5] = item.Contractor;
                    ExcelApp.Cells[i, 6] = item.Quantity;
                    ExcelApp.Cells[i, 7] = item.Procent;
                    ExcelApp.Cells[i, 8] = item.Payout;
                    i++;
                }
            }

            i = 83;
            foreach (var item in inst)
            {
                if (item.Groupe == "INSTALLATION")
                {
                    ExcelApp.Cells[i, 1] = item.Item;
                    ExcelApp.Cells[i, 2] = item.Description;
                    ExcelApp.Cells[i, 5] = item.Contractor;
                    ExcelApp.Cells[i, 6] = item.Quantity;
                    ExcelApp.Cells[i, 7] = item.Procent;
                    ExcelApp.Cells[i, 8] = item.Payout;
                    i++;
                }
            }

            i = 92;
            foreach (var item in inst)
            {
                if (item.Groupe == "OPTIONAL SERVICES")
                {
                    ExcelApp.Cells[i, 1] = item.Item;
                    ExcelApp.Cells[i, 2] = item.Description;
                    ExcelApp.Cells[i, 5] = item.Contractor;
                    ExcelApp.Cells[i, 6] = item.Quantity;
                    ExcelApp.Cells[i, 7] = item.Procent;
                    ExcelApp.Cells[i, 8] = item.Payout;
                    i++;
                }
            }

            i = 101;
            foreach (var item in cont)
            {
                ExcelApp.Cells[i, 1] = item.Contractor;
                ExcelApp.Cells[i, 2] = item.Payout;
                ExcelApp.Cells[i, 3] = item.Adjust;
                ExcelApp.Cells[i, 4] = item.Total;
                ExcelApp.Cells[i, 5] = item.TAX;
                ExcelApp.Cells[i, 6] = item.GST;
                ExcelApp.Cells[i, 7] = item.TotalContractor;
                i++;
            }

            ExcelApp.Calculate();
            ExcelApp.Cells[3, 12] = cont.FirstOrDefault().Contractor;
            ExcelApp.Cells[1, 12] = "1";   // Записуємо дані в .pdf                       
            ExcelApp.Calculate();

        }
        /// <summary>
        /// Формує файл Excel з шаблону CMO
        /// </summary>
        /// <param name="path"></param>
        private void QuotaPrintToExcelCMO(string path)
        {
            if (QuotationSelect != null && NameQuotaSelect != null)
            {
                if (NameQuotaSelect == "INVOICE")
                {
                    // Додати зміну в назві поля !!!! наприклад РІ1004
                    QuotationSelect.PrefixNumberQuota = "PI";
                }
                else
                {
                    QuotationSelect.PrefixNumberQuota = "Q";
                }


                int i = 0;
                var clientData = Clients.Where(c => c.Id == QuotationSelect.ClientId).FirstOrDefault();
                var materialData = db.MaterialQuotations.Where(q => q.QuotationId == QuotationSelect.Id);

                Excel.Application ExcelApp = new Excel.Application();
                Excel.Workbook ExcelWorkBook;
                ExcelWorkBook = ExcelApp.Workbooks.Open(Environment.CurrentDirectory + path);   //Вказуємо шлях до шаблону


                ExcelApp.Cells[1, 3] = NameQuotaSelect;

                ExcelApp.Cells[2, 5] = QuotationSelect.QuotaDate;
                ExcelApp.Cells[3, 5] = QuotationSelect.NumberQuota;
                ExcelApp.Cells[4, 5] = clientData.NumberClient;
                ExcelApp.Cells[6, 2] = QuotationSelect.JobDescription;

                ExcelApp.Cells[7, 2] = clientData.CompanyName;

                ExcelApp.Cells[8, 2] = clientData.PrimaryFirstName + " " + clientData.PrimaryLastName;
                ExcelApp.Cells[9, 2] = clientData.PrimaryPhoneNumber;
                ExcelApp.Cells[10, 2] = clientData.PrimaryEmail;
                ExcelApp.Cells[11, 2] = clientData.AddressBillStreet + ", " + clientData.AddressBillCity + ", " + clientData.AddressBillProvince + ", " + clientData.AddressBillPostalCode + ", " + clientData.AddressBillCountry;

                ExcelApp.Cells[8, 4] = clientData.SecondaryFirstName + " " + clientData.SecondaryLastName;
                ExcelApp.Cells[9, 4] = clientData.SecondaryPhoneNumber;
                ExcelApp.Cells[10, 4] = clientData.SecondaryEmail;
                ExcelApp.Cells[11, 4] = clientData.AddressSiteStreet + ", " + clientData.AddressSiteCity + ", " + clientData.AddressSiteProvince + ", " + clientData.AddressSitePostalCode + ", " + clientData.AddressSiteCountry;

                i = 19; // "FLOORING"
                foreach (var item in materialData)
                {
                    if (item.Groupe == "FLOORING")
                    {
                        //ExcelApp.Cells[i, 1] = item.Item;
                        ExcelApp.Cells[i, 1] = item.Description;
                        ExcelApp.Cells[i, 3] = item.Quantity;
                        ExcelApp.Cells[i, 4] = item.Rate;
                        ExcelApp.Cells[i, 5] = item.Price;
                        i++;
                    }
                }

                i = 28;  // "ACCESSORIES"
                foreach (var item in materialData)
                {
                    if (item.Groupe == "ACCESSORIES")
                    {
                        //ExcelApp.Cells[27 + i, 1] = item.Item;
                        ExcelApp.Cells[i, 1] = item.Description;
                        ExcelApp.Cells[i, 3] = item.Quantity;
                        ExcelApp.Cells[i, 4] = item.Rate;
                        ExcelApp.Cells[i, 5] = item.Price;
                        i++;
                    }
                }

                i = 57; // "INSTALLATION"
                foreach (var item in materialData)
                {
                    if (item.Groupe == "INSTALLATION")
                    {
                        ExcelApp.Cells[i, 1] = item.Item;
                        ExcelApp.Cells[i, 2] = item.Description;
                        ExcelApp.Cells[i, 3] = item.Quantity;
                        ExcelApp.Cells[i, 4] = item.Rate;
                        ExcelApp.Cells[i, 5] = item.Price;
                        i++;
                    }
                }

                i = 66; // "DEMOLITION"
                foreach (var item in materialData)
                {
                    if (item.Groupe == "DEMOLITION")
                    {
                        ExcelApp.Cells[i, 1] = item.Item;
                        ExcelApp.Cells[i, 2] = item.Description;
                        ExcelApp.Cells[i, 3] = item.Quantity;
                        ExcelApp.Cells[i, 4] = item.Rate;
                        ExcelApp.Cells[i, 5] = item.Price;
                        i++;
                    }
                }

                i = 72;  // "OPTIONAL SERVICES"
                foreach (var item in materialData)
                {
                    if (item.Groupe == "OPTIONAL SERVICES")
                    {
                        ExcelApp.Cells[i, 1] = item.Item;
                        ExcelApp.Cells[i, 2] = item.Description;
                        ExcelApp.Cells[i, 3] = item.Quantity;
                        ExcelApp.Cells[i, 4] = item.Rate;
                        ExcelApp.Cells[i, 5] = item.Price;
                        i++;
                    }
                }

                i = 80;  // "FLOORING DELIVERY"
                foreach (var item in materialData)
                {
                    if (item.Groupe == "FLOORING DELIVERY")
                    {
                        ExcelApp.Cells[i, 1] = item.Item;
                        ExcelApp.Cells[i, 2] = item.Description;
                        ExcelApp.Cells[i, 3] = item.Quantity;
                        ExcelApp.Cells[i, 4] = item.Rate;
                        ExcelApp.Cells[i, 5] = item.Price;
                        i++;
                    }
                }

                ExcelApp.Cells[14, 2] = QuotationSelect.JobNote;

                ExcelApp.Cells[45, 5] = QuotationSelect.MaterialSubtotal;
                ExcelApp.Cells[46, 5] = QuotationSelect.MaterialTax;
                ExcelApp.Cells[47, 5] = QuotationSelect.MaterialDiscountAmount;
                ExcelApp.Cells[48, 5] = QuotationSelect.MaterialTotal;

                ExcelApp.Cells[84, 5] = QuotationSelect.LabourSubtotal;
                ExcelApp.Cells[85, 5] = QuotationSelect.LabourTax;
                ExcelApp.Cells[86, 5] = QuotationSelect.LabourDiscountAmount;
                ExcelApp.Cells[87, 5] = QuotationSelect.LabourTotal;

                ExcelApp.Cells[90, 2] = (QuotationSelect.FinancingYesNo) ? "Yes" : "No";
                ExcelApp.Cells[91, 2] = QuotationSelect.FinancingAmount;
                ExcelApp.Cells[92, 2] = QuotationSelect.AmountPaidByCreditCard;

                ExcelApp.Cells[90, 5] = QuotationSelect.ProjectTotal;
                ExcelApp.Cells[91, 5] = QuotationSelect.FinancingFee;
                ExcelApp.Cells[92, 5] = QuotationSelect.ProcessingFee;
                ExcelApp.Cells[93, 5] = QuotationSelect.InvoiceGrandTotal;

                var pay = db.Payments.Where(p => p.QuotationId == QuotationSelect.Id).OrderBy(p => p.Id);

                int payCounter = 98;
                foreach (var item in pay)
                {
                    if (item.PaymentDatePaid == DateTime.MinValue)
                    {
                        ExcelApp.Cells[payCounter, 2] = "";
                    }
                    else
                    {
                        ExcelApp.Cells[payCounter, 2] = item.PaymentDatePaid;
                    }
                    ExcelApp.Cells[payCounter, 3] = item.PaymentAmountPaid;
                    ExcelApp.Cells[payCounter, 4] = item.PaymentMethod;

                    if (item.Balance > 0)
                    {
                        ExcelApp.Cells[payCounter, 5] = item.Balance;
                    }
                    else
                    {
                        ExcelApp.Cells[payCounter, 5] = "Paid in full";
                    }

                    payCounter++;
                }

                ExcelApp.Calculate();
                ExcelApp.Cells[1, 9] = "1";   // Записуємо дані в .pdf    
                ExcelApp.Calculate();

                db.Entry(QuotationSelect).State = EntityState.Modified;

                var delivery = Deliveries.FirstOrDefault(d => d.QuotaId == QuotationSelect.Id);
                var workOrder = WorkOrders.FirstOrDefault(w => w.QuotaId == QuotationSelect.Id);
                var invoice = Invoices.FirstOrDefault(inv => inv.QuotaId == QuotationSelect.Id);

                if (delivery != null)
                {
                    delivery.NumberQuota = QuotationSelect.NumberQuota;
                    db.Entry(delivery).State = EntityState.Modified;
                }
                if (workOrder != null)
                {
                    workOrder.NumberQuota = QuotationSelect.NumberQuota;
                    db.Entry(workOrder).State = EntityState.Modified;
                }

                db.SaveChanges();
                Deliveries = null;
                WorkOrders = null;
                Quotations = null;

                LoadQuotationsDB(CompanyName);
                LoadWorkOrdersDB(CompanyName);
                LoadDeliveriesDB(CompanyName, false);
            }
        }
        /// <summary>
        /// Формує файл Excel з шаблону Next Level
        /// </summary>
        /// <param name="path"></param>
        private void QuotaPrintToExcelNL(string path)
        {
            if (QuotationSelect != null && NameQuotaSelect != null)
            {
                if (NameQuotaSelect == "INVOICE")
                {
                    // Додати зміну в назві поля !!!! наприклад РІ1004
                    QuotationSelect.PrefixNumberQuota = "PI";
                }
                else
                {
                    QuotationSelect.PrefixNumberQuota = "Q";
                }


                int i = 0;
                var clientData = Clients.Where(c => c.Id == QuotationSelect.ClientId).FirstOrDefault();
                var materialData = db.MaterialQuotations.Where(q => q.QuotationId == QuotationSelect.Id);

                Excel.Application ExcelApp = new Excel.Application();
                Excel.Workbook ExcelWorkBook;
                ExcelWorkBook = ExcelApp.Workbooks.Open(Environment.CurrentDirectory + path);   //Вказуємо шлях до шаблону


                ExcelApp.Cells[1, 5] = NameQuotaSelect;

                ExcelApp.Cells[2, 7] = QuotationSelect.QuotaDate;
                ExcelApp.Cells[3, 7] = QuotationSelect.NumberQuota;
                ExcelApp.Cells[4, 7] = clientData.NumberClient;
                ExcelApp.Cells[6, 2] = QuotationSelect.JobDescription;

                ExcelApp.Cells[7, 2] = clientData.CompanyName;

                ExcelApp.Cells[8, 2] = clientData.PrimaryFirstName + " " + clientData.PrimaryLastName;
                ExcelApp.Cells[9, 2] = clientData.PrimaryPhoneNumber;
                ExcelApp.Cells[10, 2] = clientData.PrimaryEmail;
                ExcelApp.Cells[11, 2] = clientData.AddressBillStreet + ", " + clientData.AddressBillCity + ", " + clientData.AddressBillProvince + ", " + clientData.AddressBillPostalCode + ", " + clientData.AddressBillCountry;

                ExcelApp.Cells[8, 5] = clientData.SecondaryFirstName + " " + clientData.SecondaryLastName;
                ExcelApp.Cells[9, 5] = clientData.SecondaryPhoneNumber;
                ExcelApp.Cells[10, 5] = clientData.SecondaryEmail;
                ExcelApp.Cells[11, 5] = clientData.AddressSiteStreet + ", " + clientData.AddressSiteCity + ", " + clientData.AddressSiteProvince + ", " + clientData.AddressSitePostalCode + ", " + clientData.AddressSiteCountry;

                ExcelApp.Cells[14, 2] = QuotationSelect.JobNote;

                i = 19; // "FLOORING"
                foreach (var item in materialData)
                {
                    if (item.Groupe == "FLOORING")
                    {
                        //ExcelApp.Cells[i, 1] = item.Item;
                        ExcelApp.Cells[i, 1] = item.Description;
                        ExcelApp.Cells[i, 3] = item.QuantityNL;
                        ExcelApp.Cells[i, 4] = item.Depth;
                        ExcelApp.Cells[i, 5] = item.Mapei;
                        ExcelApp.Cells[i, 6] = item.Rate;
                        ExcelApp.Cells[i, 7] = item.Price;
                        i++;
                    }
                }

                i = 28;  // "ACCESSORIES"
                foreach (var item in materialData)
                {
                    if (item.Groupe == "ACCESSORIES")
                    {
                        //ExcelApp.Cells[27 + i, 1] = item.Item;
                        ExcelApp.Cells[i, 1] = item.Description;
                        ExcelApp.Cells[i, 5] = item.Quantity;
                        ExcelApp.Cells[i, 6] = item.Rate;
                        ExcelApp.Cells[i, 7] = item.Price;
                        i++;
                    }
                }

                i = 57; // "INSTALLATION"
                foreach (var item in materialData)
                {
                    if (item.Groupe == "INSTALLATION")
                    {
                        ExcelApp.Cells[i, 1] = item.Item;
                        ExcelApp.Cells[i, 2] = item.Description;
                        ExcelApp.Cells[i, 5] = item.Quantity;
                        ExcelApp.Cells[i, 6] = item.Rate;
                        ExcelApp.Cells[i, 7] = item.Price;
                        i++;
                    }
                }

                i = 66; // "DEMOLITION"
                foreach (var item in materialData)
                {
                    if (item.Groupe == "DEMOLITION")
                    {
                        ExcelApp.Cells[i, 1] = item.Item;
                        ExcelApp.Cells[i, 2] = item.Description;
                        ExcelApp.Cells[i, 5] = item.Quantity;
                        ExcelApp.Cells[i, 6] = item.Rate;
                        ExcelApp.Cells[i, 7] = item.Price;
                        i++;
                    }
                }

                i = 72;  // "OPTIONAL SERVICES"
                foreach (var item in materialData)
                {
                    if (item.Groupe == "OPTIONAL SERVICES")
                    {
                        ExcelApp.Cells[i, 1] = item.Item;
                        ExcelApp.Cells[i, 2] = item.Description;
                        ExcelApp.Cells[i, 5] = item.Quantity;
                        ExcelApp.Cells[i, 6] = item.Rate;
                        ExcelApp.Cells[i, 7] = item.Price;
                        i++;
                    }
                }

                i = 80;  // "FLOORING DELIVERY"
                foreach (var item in materialData)
                {
                    if (item.Groupe == "FLOORING DELIVERY")
                    {
                        ExcelApp.Cells[i, 1] = item.Item;
                        ExcelApp.Cells[i, 2] = item.Description;
                        ExcelApp.Cells[i, 5] = item.Quantity;
                        ExcelApp.Cells[i, 6] = item.Rate;
                        ExcelApp.Cells[i, 7] = item.Price;
                        i++;
                    }
                }


                ExcelApp.Cells[45, 7] = QuotationSelect.MaterialSubtotal;
                ExcelApp.Cells[46, 7] = QuotationSelect.MaterialTax;
                ExcelApp.Cells[47, 7] = QuotationSelect.MaterialDiscountAmount;
                ExcelApp.Cells[48, 7] = QuotationSelect.MaterialTotal;

                ExcelApp.Cells[84, 7] = QuotationSelect.LabourSubtotal;
                ExcelApp.Cells[85, 7] = QuotationSelect.LabourTax;
                ExcelApp.Cells[86, 7] = QuotationSelect.LabourDiscountAmount;
                ExcelApp.Cells[87, 7] = QuotationSelect.LabourTotal;

                ExcelApp.Cells[90, 2] = (QuotationSelect.FinancingYesNo) ? "Yes" : "No";
                ExcelApp.Cells[91, 2] = QuotationSelect.FinancingAmount;
                ExcelApp.Cells[92, 2] = QuotationSelect.AmountPaidByCreditCard;

                ExcelApp.Cells[90, 7] = QuotationSelect.ProjectTotal;
                ExcelApp.Cells[91, 7] = QuotationSelect.FinancingFee;
                ExcelApp.Cells[92, 7] = QuotationSelect.ProcessingFee;
                ExcelApp.Cells[93, 7] = QuotationSelect.InvoiceGrandTotal;

                var pay = db.Payments.Where(p => p.QuotationId == QuotationSelect.Id).OrderBy(p => p.Id);

                int payCounter = 98;
                foreach (var item in pay)
                {
                    if (item.PaymentDatePaid == DateTime.MinValue)
                    {
                        ExcelApp.Cells[payCounter, 2] = "";
                    }
                    else
                    {
                        ExcelApp.Cells[payCounter, 2] = item.PaymentDatePaid;
                    }
                    ExcelApp.Cells[payCounter, 5] = item.PaymentAmountPaid;
                    ExcelApp.Cells[payCounter, 6] = item.PaymentMethod;

                    if (item.Balance > 0)
                    {
                        ExcelApp.Cells[payCounter, 7] = item.Balance;
                    }
                    else
                    {
                        ExcelApp.Cells[payCounter, 7] = "Paid in full";
                    }

                    payCounter++;
                }

                ExcelApp.Calculate();
                ExcelApp.Cells[1, 9] = "1";   // Записуємо дані в .pdf    
                ExcelApp.Calculate();

                db.Entry(QuotationSelect).State = EntityState.Modified;

                var delivery = Deliveries.FirstOrDefault(d => d.QuotaId == QuotationSelect.Id);
                var workOrder = WorkOrders.FirstOrDefault(w => w.QuotaId == QuotationSelect.Id);
                var invoice = Invoices.FirstOrDefault(inv => inv.QuotaId == QuotationSelect.Id);

                if (delivery != null)
                {
                    delivery.NumberQuota = QuotationSelect.NumberQuota;
                    db.Entry(delivery).State = EntityState.Modified;
                }
                if (workOrder != null)
                {
                    workOrder.NumberQuota = QuotationSelect.NumberQuota;
                    db.Entry(workOrder).State = EntityState.Modified;
                }

                db.SaveChanges();
                Deliveries = null;
                WorkOrders = null;
                Quotations = null;

                LoadQuotationsDB(CompanyName);
                LoadWorkOrdersDB(CompanyName);
                LoadDeliveriesDB(CompanyName, false);
            }
        }
        /// <summary>
        /// Формує файл Excel з шаблону CMO
        /// </summary>
        /// <param name="path"></param>
        private void DeliveryPrintToExcelCMO(string path)
        {
            if (DeliverySelect != null)
            {
                var material = db.DeliveryMaterials.Where(m => m.DeliveryId == DeliverySelect.Id);
                var dic = db.DIC_Suppliers.FirstOrDefault(s => s.Id == DeliverySelect.SupplierId);
                var quota = Quotations.FirstOrDefault(q => q.Id == DeliverySelect.QuotaId); //db.Quotations.FirstOrDefault(q => q.Id == DeliverySelect.QuotaId);
                var client = Clients.FirstOrDefault(c => c.Id == quota.ClientId); //db.Clients.FirstOrDefault(c => c.Id == quota.ClientId);
                Excel.Application ExcelApp = new Excel.Application();
                Excel.Workbook ExcelWorkBook;
                ExcelWorkBook = ExcelApp.Workbooks.Open(Environment.CurrentDirectory + path);   //Вказуємо шлях до шаблону

                ExcelApp.Cells[3, 6] = DeliverySelect.DateCreating;
                ExcelApp.Cells[6, 2] = "\"" + client?.NumberClient + " - " + quota?.NumberQuota + "\"";
                ExcelApp.Cells[7, 2] = "\"CMO" + " - " + client?.PrimaryFullName + "\"";
                ExcelApp.Cells[9, 2] = dic?.Supplier;
                ExcelApp.Cells[10, 2] = dic?.Address;

                int i = 14;
                foreach (var item in material)
                {
                    ExcelApp.Cells[i, 1] = item.Description;
                    ExcelApp.Cells[i, 6] = item.Quantity;
                    i++;
                }

                ExcelApp.Cells[1, 9] = "1";   // Записуємо дані в .pdf    
                ExcelApp.Calculate();

                DeliverySelect.Color = "Black";
                DeliverySelect.IsEnabled = true;
                db.Entry(DeliverySelect).State = EntityState.Modified;
                db.SaveChanges();

                Deliveries = null;
                LoadDeliveriesDB(CompanyName, false);
            }
        }
        /// <summary>
        /// Формує файл Excel з шаблону Next Level
        /// </summary>
        /// <param name="path"></param>
        private void DeliveryPrintToExcelNL(string path)
        {
            if (DeliverySelect != null)
            {
                var material = db.DeliveryMaterials.Where(m => m.DeliveryId == DeliverySelect.Id);
                var dic = db.DIC_Suppliers.FirstOrDefault(s => s.Id == DeliverySelect.SupplierId);
                var quota = Quotations.FirstOrDefault(q => q.Id == DeliverySelect.QuotaId); //db.Quotations.FirstOrDefault(q => q.Id == DeliverySelect.QuotaId);
                var client = Clients.FirstOrDefault(c => c.Id == quota.ClientId); //db.Clients.FirstOrDefault(c => c.Id == quota.ClientId);
                Excel.Application ExcelApp = new Excel.Application();
                Excel.Workbook ExcelWorkBook;
                ExcelWorkBook = ExcelApp.Workbooks.Open(Environment.CurrentDirectory + path);   //Вказуємо шлях до шаблону

                ExcelApp.Cells[3, 6] = DeliverySelect.DateCreating;
                ExcelApp.Cells[6, 2] = "\"" + client?.NumberClient + " - " + quota?.NumberQuota + "\"";
                ExcelApp.Cells[7, 2] = "\"CMO" + " - " + client?.PrimaryFullName + "\"";
                ExcelApp.Cells[9, 2] = dic?.Supplier;
                ExcelApp.Cells[10, 2] = dic?.Address;

                int i = 14;
                foreach (var item in material)
                {
                    ExcelApp.Cells[i, 1] = item.Description;
                    ExcelApp.Cells[i, 6] = item.Quantity;
                    i++;
                }

                ExcelApp.Cells[1, 9] = "1";   // Записуємо дані в .pdf    
                ExcelApp.Calculate();

                DeliverySelect.Color = "Black";
                DeliverySelect.IsEnabled = true;
                db.Entry(DeliverySelect).State = EntityState.Modified;
                db.SaveChanges();

                Deliveries = null;
                LoadDeliveriesDB(CompanyName, false);
            }
        }
        /// <summary>
        /// /// Формує файл Excel з шаблону CMO
        /// </summary>
        /// <param name="path"></param>
        private void DeliveryDriverPrintYoExcelCMO(string path)
        {
            if (DeliveryComboBoxSelect != null)
            {
                Excel.Application ExcelApp = new Excel.Application();
                Excel.Workbook ExcelWorkBook;
                ExcelWorkBook = ExcelApp.Workbooks.Open(Environment.CurrentDirectory + path);   //Вказуємо шлях до шаблону

                string[] comboBox = DeliveryComboBoxSelect.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                string word = comboBox[0];
                var supplier = db.Deliveries.Where(d => d.NumberQuota == word);
                int? quotaId = supplier.FirstOrDefault(s => s.QuotaId > 0)?.QuotaId;
                var quota = Quotations.FirstOrDefault(q => q.Id == quotaId); //db.Quotations.FirstOrDefault(q => q.Id == quotaId);
                var client = Clients.FirstOrDefault(c => c.Id == quota.ClientId); //db.Clients.FirstOrDefault(c => c.Id == quota.ClientId);
                if (supplier != null && quota != null && client != null)
                {
                    ExcelApp.Cells[3, 5] = DateTime.Today;
                    ExcelApp.Cells[4, 5] = quota.NumberQuota;
                    ExcelApp.Cells[9, 2] = client.PrimaryFullName;
                    ExcelApp.Cells[10, 2] = client?.CompanyName;
                    ExcelApp.Cells[11, 2] = client?.AddressBillStreet + client?.AddressBillCity + client?.AddressBillProvince + client?.AddressBillPostalCode + client?.AddressBillCountry;
                    ExcelApp.Cells[12, 2] = client?.PrimaryPhoneNumber;
                    ExcelApp.Cells[11, 4] = client?.AddressSiteStreet + client?.AddressSiteCity + client?.AddressSiteProvince + client?.AddressSitePostalCode + client?.AddressSiteCountry;
                    ExcelApp.Cells[12, 4] = client?.PrimaryEmail;
                    int i = 24;
                    foreach (var item in supplier)
                    {
                        var temp = db.DeliveryMaterials.Where(m => m.DeliveryId == item.Id)?.Select(m => m.Quantity);
                        decimal quantity = temp?.Sum() ?? (0m);
                        var dictionary = db.DIC_Suppliers.FirstOrDefault(s => s.Id == item.SupplierId);

                        ExcelApp.Cells[i, 1] = dictionary.Supplier;
                        ExcelApp.Cells[i, 2] = dictionary.Address;
                        ExcelApp.Cells[i, 3] = dictionary.Hours;
                        ExcelApp.Cells[i, 4] = quantity;
                        ExcelApp.Cells[i, 5] = item.OrderNumber;
                        i++;
                    }
                }
                ExcelApp.Cells[1, 9] = "1";   // Записуємо дані в .pdf    
                ExcelApp.Calculate();
            }
        }
        /// <summary>
        /// Формує файл Excel з шаблону Next Level
        /// </summary>
        /// <param name="path"></param>
        private void DeliveryDriverPrintYoExcelNL(string path)
        {
            if (DeliveryComboBoxSelect != null)
            {
                Excel.Application ExcelApp = new Excel.Application();
                Excel.Workbook ExcelWorkBook;
                ExcelWorkBook = ExcelApp.Workbooks.Open(Environment.CurrentDirectory + path);   //Вказуємо шлях до шаблону

                string[] comboBox = DeliveryComboBoxSelect.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                string word = comboBox[0];
                var supplier = db.Deliveries.Where(d => d.NumberQuota == word);
                int? quotaId = supplier.FirstOrDefault(s => s.QuotaId > 0)?.QuotaId;
                var quota = Quotations.FirstOrDefault(q => q.Id == quotaId); //db.Quotations.FirstOrDefault(q => q.Id == quotaId);
                var client = Clients.FirstOrDefault(c => c.Id == quota.ClientId); //db.Clients.FirstOrDefault(c => c.Id == quota.ClientId);
                if (supplier != null && quota != null && client != null)
                {
                    ExcelApp.Cells[3, 5] = DateTime.Today;
                    ExcelApp.Cells[4, 5] = quota.NumberQuota;
                    ExcelApp.Cells[9, 2] = client.PrimaryFullName;
                    ExcelApp.Cells[10, 2] = client?.CompanyName;
                    ExcelApp.Cells[11, 2] = client?.AddressBillStreet + client?.AddressBillCity + client?.AddressBillProvince + client?.AddressBillPostalCode + client?.AddressBillCountry;
                    ExcelApp.Cells[12, 2] = client?.PrimaryPhoneNumber;
                    ExcelApp.Cells[11, 4] = client?.AddressSiteStreet + client?.AddressSiteCity + client?.AddressSiteProvince + client?.AddressSitePostalCode + client?.AddressSiteCountry;
                    ExcelApp.Cells[12, 4] = client?.PrimaryEmail;
                    int i = 24;
                    foreach (var item in supplier)
                    {
                        var temp = db.DeliveryMaterials.Where(m => m.DeliveryId == item.Id)?.Select(m => m.Quantity);
                        decimal quantity = temp?.Sum() ?? (0m);
                        var dictionary = db.DIC_Suppliers.FirstOrDefault(s => s.Id == item.SupplierId);

                        ExcelApp.Cells[i, 1] = dictionary.Supplier;
                        ExcelApp.Cells[i, 2] = dictionary.Address;
                        ExcelApp.Cells[i, 3] = dictionary.Hours;
                        ExcelApp.Cells[i, 4] = quantity;
                        ExcelApp.Cells[i, 5] = item.OrderNumber;
                        i++;
                    }
                }
                ExcelApp.Cells[1, 9] = "1";   // Записуємо дані в .pdf    
                ExcelApp.Calculate();
            }
        }
        /// <summary>
        /// /// Формує файл Excel з шаблону CMO
        /// </summary>
        /// <param name="path"></param>
        private void InvoicesPrintToExcelCMO(string path)
        {
            if (InvoiceSelect != null)
            {
                int i = 0;
                var quota = Quotations.FirstOrDefault(q => q.Id == InvoiceSelect.QuotaId); //db.Quotations.FirstOrDefault(q => q.Id == InvoiceSelect.QuotaId);
                var clientData = Clients.Where(c => c.Id == quota.ClientId).FirstOrDefault(); //db.Clients.Where(c => c.Id == quota.ClientId).FirstOrDefault();
                var materialData = db.MaterialQuotations.Where(q => q.QuotationId == quota.Id);

                Excel.Application ExcelApp = new Excel.Application();
                Excel.Workbook ExcelWorkBook;
                ExcelWorkBook = ExcelApp.Workbooks.Open(Environment.CurrentDirectory + path);   //Вказуємо шлях до шаблону

                //ExcelApp.Cells[1, 3] = NameQuotaSelect;
                ExcelApp.Cells[2, 3] = clientData.NumberClient;
                ExcelApp.Cells[3, 3] = InvoiceSelect.UpNumber;
                ExcelApp.Cells[4, 3] = InvoiceSelect.OrderNumber;

                ExcelApp.Cells[2, 5] = InvoiceSelect.DateInvoice;
                ExcelApp.Cells[3, 5] = InvoiceSelect.NumberInvoice;
                ExcelApp.Cells[4, 5] = quota.NumberQuota;

                ExcelApp.Cells[6, 2] = quota.JobDescription;

                ExcelApp.Cells[7, 2] = clientData.CompanyName;

                ExcelApp.Cells[8, 2] = clientData.PrimaryFirstName + " " + clientData.PrimaryLastName;
                ExcelApp.Cells[9, 2] = clientData.PrimaryPhoneNumber;
                ExcelApp.Cells[10, 2] = clientData.PrimaryEmail;
                ExcelApp.Cells[11, 2] = clientData.AddressBillStreet + ", " + clientData.AddressBillCity + ", " + clientData.AddressBillProvince + ", " + clientData.AddressBillPostalCode + ", " + clientData.AddressBillCountry;

                ExcelApp.Cells[8, 4] = clientData.SecondaryFirstName + " " + clientData.SecondaryLastName;
                ExcelApp.Cells[9, 4] = clientData.SecondaryPhoneNumber;
                ExcelApp.Cells[10, 4] = clientData.SecondaryEmail;
                ExcelApp.Cells[11, 4] = clientData.AddressSiteStreet + ", " + clientData.AddressSiteCity + ", " + clientData.AddressSiteProvince + ", " + clientData.AddressSitePostalCode + ", " + clientData.AddressSiteCountry;

                i = 0;
                foreach (var item in materialData)
                {
                    if (item.Groupe == "FLOORING")
                    {
                        i++;
                        //ExcelApp.Cells[18 + i, 1] = item.Item;
                        ExcelApp.Cells[18 + i, 1] = item.Description;
                        ExcelApp.Cells[18 + i, 3] = item.Quantity;
                        ExcelApp.Cells[18 + i, 4] = item.Rate;
                        ExcelApp.Cells[18 + i, 5] = item.Price;
                    }
                }

                i = 0;
                foreach (var item in materialData)
                {
                    if (item.Groupe == "ACCESSORIES")
                    {
                        i++;
                        //ExcelApp.Cells[27 + i, 1] = item.Item;
                        ExcelApp.Cells[27 + i, 1] = item.Description;
                        ExcelApp.Cells[27 + i, 3] = item.Quantity;
                        ExcelApp.Cells[27 + i, 4] = item.Rate;
                        ExcelApp.Cells[27 + i, 5] = item.Price;
                    }
                }

                i = 0;
                foreach (var item in materialData)
                {
                    if (item.Groupe == "INSTALLATION")
                    {
                        i++;
                        ExcelApp.Cells[56 + i, 1] = item.Item;
                        ExcelApp.Cells[56 + i, 2] = item.Description;
                        ExcelApp.Cells[56 + i, 3] = item.Quantity;
                        ExcelApp.Cells[56 + i, 4] = item.Rate;
                        ExcelApp.Cells[56 + i, 5] = item.Price;
                    }
                }

                i = 0;
                foreach (var item in materialData)
                {
                    if (item.Groupe == "DEMOLITION")
                    {
                        i++;
                        ExcelApp.Cells[65 + i, 1] = item.Item;
                        ExcelApp.Cells[65 + i, 2] = item.Description;
                        ExcelApp.Cells[65 + i, 3] = item.Quantity;
                        ExcelApp.Cells[65 + i, 4] = item.Rate;
                        ExcelApp.Cells[65 + i, 5] = item.Price;
                    }
                }

                i = 0;
                foreach (var item in materialData)
                {
                    if (item.Groupe == "OPTIONAL SERVICES")
                    {
                        i++;
                        ExcelApp.Cells[71 + i, 1] = item.Item;
                        ExcelApp.Cells[71 + i, 2] = item.Description;
                        ExcelApp.Cells[71 + i, 3] = item.Quantity;
                        ExcelApp.Cells[71 + i, 4] = item.Rate;
                        ExcelApp.Cells[71 + i, 5] = item.Price;
                    }
                }

                i = 0;
                foreach (var item in materialData)
                {
                    if (item.Groupe == "FLOORING DELIVERY")
                    {
                        i++;
                        ExcelApp.Cells[79 + i, 1] = item.Item;
                        ExcelApp.Cells[79 + i, 2] = item.Description;
                        ExcelApp.Cells[79 + i, 3] = item.Quantity;
                        ExcelApp.Cells[79 + i, 4] = item.Rate;
                        ExcelApp.Cells[79 + i, 5] = item.Price;
                    }
                }

                ExcelApp.Cells[14, 2] = quota.JobNote;

                ExcelApp.Cells[45, 5] = quota.MaterialSubtotal;
                ExcelApp.Cells[46, 5] = quota.MaterialTax;
                ExcelApp.Cells[47, 5] = quota.MaterialDiscountAmount;
                ExcelApp.Cells[48, 5] = quota.MaterialTotal;

                ExcelApp.Cells[84, 5] = quota.LabourSubtotal;
                ExcelApp.Cells[85, 5] = quota.LabourTax;
                ExcelApp.Cells[86, 5] = quota.LabourDiscountAmount;
                ExcelApp.Cells[87, 5] = quota.LabourTotal;

                ExcelApp.Cells[90, 2] = (quota.FinancingYesNo) ? "Yes" : "No";
                ExcelApp.Cells[91, 2] = quota.FinancingAmount;
                ExcelApp.Cells[92, 2] = quota.AmountPaidByCreditCard;

                ExcelApp.Cells[90, 5] = quota.ProjectTotal;
                ExcelApp.Cells[91, 5] = quota.FinancingFee;
                ExcelApp.Cells[92, 5] = quota.ProcessingFee;
                ExcelApp.Cells[93, 5] = quota.InvoiceGrandTotal;

                var pay = db.Payments.Where(p => p.QuotationId == quota.Id).OrderBy(p => p.Id);

                int payCounter = 98;
                foreach (var item in pay)
                {
                    if (item.PaymentDatePaid == DateTime.MinValue)
                    {
                        ExcelApp.Cells[payCounter, 2] = "";
                    }
                    else
                    {
                        ExcelApp.Cells[payCounter, 2] = item.PaymentDatePaid;
                    }
                    ExcelApp.Cells[payCounter, 3] = item.PaymentAmountPaid;
                    ExcelApp.Cells[payCounter, 4] = item.PaymentMethod;

                    if (item.Balance > 0)
                    {
                        ExcelApp.Cells[payCounter, 5] = item.Balance;
                    }
                    else
                    {
                        ExcelApp.Cells[payCounter, 5] = "Paid in full";
                    }

                    payCounter++;
                }

                ExcelApp.Calculate();
                ExcelApp.Cells[1, 9] = "1";   // Записуємо дані в .pdf    
                ExcelApp.Calculate();

            }

        }
        /// <summary>
        /// Формує файл Excel з шаблону Next Level
        /// </summary>
        /// <param name="path"></param>
        private void InvoicePrintToExcelNL(string path)
        {
            if (InvoiceSelect != null)
            {
                int i = 0;
                var quota = Quotations.FirstOrDefault(q => q.Id == InvoiceSelect.QuotaId);
                var clientData = Clients.Where(c => c.Id == quota.ClientId).FirstOrDefault();
                var materialData = db.MaterialQuotations.Where(q => q.QuotationId == quota.Id);

                Excel.Application ExcelApp = new Excel.Application();
                Excel.Workbook ExcelWorkBook;
                ExcelWorkBook = ExcelApp.Workbooks.Open(Environment.CurrentDirectory + path);   //Вказуємо шлях до шаблону


                //ExcelApp.Cells[1, 5] = //NameQuotaSelect;

                ExcelApp.Cells[2, 7] = InvoiceSelect.DateInvoice;
                ExcelApp.Cells[3, 7] = InvoiceSelect.NumberQuota;
                ExcelApp.Cells[4, 7] = clientData.NumberClient;
                ExcelApp.Cells[5, 7] = InvoiceSelect.NumberInvoice;
                ExcelApp.Cells[6, 2] = quota.JobDescription;

                ExcelApp.Cells[7, 2] = clientData.CompanyName;

                ExcelApp.Cells[8, 2] = clientData.PrimaryFirstName + " " + clientData.PrimaryLastName;
                ExcelApp.Cells[9, 2] = clientData.PrimaryPhoneNumber;
                ExcelApp.Cells[10, 2] = clientData.PrimaryEmail;
                ExcelApp.Cells[11, 2] = clientData.AddressBillStreet + ", " + clientData.AddressBillCity + ", " + clientData.AddressBillProvince + ", " + clientData.AddressBillPostalCode + ", " + clientData.AddressBillCountry;

                ExcelApp.Cells[8, 5] = clientData.SecondaryFirstName + " " + clientData.SecondaryLastName;
                ExcelApp.Cells[9, 5] = clientData.SecondaryPhoneNumber;
                ExcelApp.Cells[10, 5] = clientData.SecondaryEmail;
                ExcelApp.Cells[11, 5] = clientData.AddressSiteStreet + ", " + clientData.AddressSiteCity + ", " + clientData.AddressSiteProvince + ", " + clientData.AddressSitePostalCode + ", " + clientData.AddressSiteCountry;

                ExcelApp.Cells[14, 2] = quota.JobNote;

                i = 19; // "FLOORING"
                foreach (var item in materialData)
                {
                    if (item.Groupe == "FLOORING")
                    {
                        //ExcelApp.Cells[i, 1] = item.Item;
                        ExcelApp.Cells[i, 1] = item.Description;
                        ExcelApp.Cells[i, 3] = item.QuantityNL;
                        ExcelApp.Cells[i, 4] = item.Depth;
                        ExcelApp.Cells[i, 5] = item.Mapei;
                        ExcelApp.Cells[i, 6] = item.Rate;
                        ExcelApp.Cells[i, 7] = item.Price;
                        i++;
                    }
                }

                i = 28;  // "ACCESSORIES"
                foreach (var item in materialData)
                {
                    if (item.Groupe == "ACCESSORIES")
                    {
                        //ExcelApp.Cells[27 + i, 1] = item.Item;
                        ExcelApp.Cells[i, 1] = item.Description;
                        ExcelApp.Cells[i, 5] = item.Quantity;
                        ExcelApp.Cells[i, 6] = item.Rate;
                        ExcelApp.Cells[i, 7] = item.Price;
                        i++;
                    }
                }

                i = 57; // "INSTALLATION"
                foreach (var item in materialData)
                {
                    if (item.Groupe == "INSTALLATION")
                    {
                        ExcelApp.Cells[i, 1] = item.Item;
                        ExcelApp.Cells[i, 2] = item.Description;
                        ExcelApp.Cells[i, 5] = item.Quantity;
                        ExcelApp.Cells[i, 6] = item.Rate;
                        ExcelApp.Cells[i, 7] = item.Price;
                        i++;
                    }
                }

                i = 66; // "DEMOLITION"
                foreach (var item in materialData)
                {
                    if (item.Groupe == "DEMOLITION")
                    {
                        ExcelApp.Cells[i, 1] = item.Item;
                        ExcelApp.Cells[i, 2] = item.Description;
                        ExcelApp.Cells[i, 5] = item.Quantity;
                        ExcelApp.Cells[i, 6] = item.Rate;
                        ExcelApp.Cells[i, 7] = item.Price;
                        i++;
                    }
                }

                i = 72;  // "OPTIONAL SERVICES"
                foreach (var item in materialData)
                {
                    if (item.Groupe == "OPTIONAL SERVICES")
                    {
                        ExcelApp.Cells[i, 1] = item.Item;
                        ExcelApp.Cells[i, 2] = item.Description;
                        ExcelApp.Cells[i, 5] = item.Quantity;
                        ExcelApp.Cells[i, 6] = item.Rate;
                        ExcelApp.Cells[i, 7] = item.Price;
                        i++;
                    }
                }

                i = 80;  // "FLOORING DELIVERY"
                foreach (var item in materialData)
                {
                    if (item.Groupe == "FLOORING DELIVERY")
                    {
                        ExcelApp.Cells[i, 1] = item.Item;
                        ExcelApp.Cells[i, 2] = item.Description;
                        ExcelApp.Cells[i, 5] = item.Quantity;
                        ExcelApp.Cells[i, 6] = item.Rate;
                        ExcelApp.Cells[i, 7] = item.Price;
                        i++;
                    }
                }


                ExcelApp.Cells[45, 7] = quota.MaterialSubtotal;
                ExcelApp.Cells[46, 7] = quota.MaterialTax;
                ExcelApp.Cells[47, 7] = quota.MaterialDiscountAmount;
                ExcelApp.Cells[48, 7] = quota.MaterialTotal;

                ExcelApp.Cells[84, 7] = quota.LabourSubtotal;
                ExcelApp.Cells[85, 7] = quota.LabourTax;
                ExcelApp.Cells[86, 7] = quota.LabourDiscountAmount;
                ExcelApp.Cells[87, 7] = quota.LabourTotal;

                ExcelApp.Cells[90, 2] = (quota.FinancingYesNo) ? "Yes" : "No";
                ExcelApp.Cells[91, 2] = quota.FinancingAmount;
                ExcelApp.Cells[92, 2] = quota.AmountPaidByCreditCard;

                ExcelApp.Cells[90, 7] = quota.ProjectTotal;
                ExcelApp.Cells[91, 7] = quota.FinancingFee;
                ExcelApp.Cells[92, 7] = quota.ProcessingFee;
                ExcelApp.Cells[93, 7] = quota.InvoiceGrandTotal;

                var pay = db.Payments.Where(p => p.QuotationId == quota.Id).OrderBy(p => p.Id);

                int payCounter = 98;
                foreach (var item in pay)
                {
                    if (item.PaymentDatePaid == DateTime.MinValue)
                    {
                        ExcelApp.Cells[payCounter, 2] = "";
                    }
                    else
                    {
                        ExcelApp.Cells[payCounter, 2] = item.PaymentDatePaid;
                    }
                    ExcelApp.Cells[payCounter, 5] = item.PaymentAmountPaid;
                    ExcelApp.Cells[payCounter, 6] = item.PaymentMethod;

                    if (item.Balance > 0)
                    {
                        ExcelApp.Cells[payCounter, 7] = item.Balance;
                    }
                    else
                    {
                        ExcelApp.Cells[payCounter, 7] = "Paid in full";
                    }

                    payCounter++;
                }

                ExcelApp.Calculate();
                ExcelApp.Cells[1, 9] = "1";   // Записуємо дані в .pdf    
                ExcelApp.Calculate();

            }
        }
    }
}
