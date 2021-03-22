using Builders.ViewModels;
using Builders.Views;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace Builders
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public DisplayRootRegistry displayRootRegistry = new DisplayRootRegistry();
        public decimal constGST { get; set; }
        public decimal constTAX { get; set; }

        public App()
        {
            displayRootRegistry.RegisterWindowType<MainViewModel, MainView>();
            displayRootRegistry.RegisterWindowType<DetailsViewModel, DetailsView>();
            displayRootRegistry.RegisterWindowType<ClientViewModel, ClientView>();
            displayRootRegistry.RegisterWindowType<DIC_ClientViewModel, DIC_ClientView>();
            displayRootRegistry.RegisterWindowType<DIC_QuotationViewModel, DIC_QuotationView>();
            displayRootRegistry.RegisterWindowType<DIC_addItemViewModel, DIC_addItemView>();
            displayRootRegistry.RegisterWindowType<DIC_addDescriptionViewModel, DIC_addDescriptionView>();
            displayRootRegistry.RegisterWindowType<DIC_ContractorViewModel, DIC_ContractorView>();
            displayRootRegistry.RegisterWindowType<DIC_SupplierViewModel, DIC_SupplierView>();
            displayRootRegistry.RegisterWindowType<DIC_DepthViewModel, DIC_DepthView>();
            displayRootRegistry.RegisterWindowType<QuotationViewModel, QuotationView>();
            displayRootRegistry.RegisterWindowType<SelectClientViewModel, SelectClientView>();
            displayRootRegistry.RegisterWindowType<SelectQuotaViewModel, SelectQuotaView>();
            displayRootRegistry.RegisterWindowType<QuotaOtherViewModel, QuotaOtherView>();
            displayRootRegistry.RegisterWindowType<PaymentViewModel, PaymentView>();
            displayRootRegistry.RegisterWindowType<InvoiceViewModel, InvoiceView>();
            displayRootRegistry.RegisterWindowType<WorkOrderViewModel, WorkOrderView>();
            displayRootRegistry.RegisterWindowType<MaterialProfitViewModel, MaterialProfitView>();
            displayRootRegistry.RegisterWindowType<LabourProfitViewModel, LabourProfitView>();
            displayRootRegistry.RegisterWindowType<FotoViewModel, FotoView>();
            displayRootRegistry.RegisterWindowType<ExpensesViewModel, ExpensesView>();
            displayRootRegistry.RegisterWindowType<DebtsViewModel, DebtsView>();
            displayRootRegistry.RegisterWindowType<DebtsPaymentViewModel, DebtsPaymentView>();
            displayRootRegistry.RegisterWindowType<MessageViewModel, MessageView>();
            displayRootRegistry.RegisterWindowType<DeliveryViewModel, DeliveryView>();
            displayRootRegistry.RegisterWindowType<GeneratedViewModel, GeneratedView>();
            displayRootRegistry.RegisterWindowType<QuotaItemEditViewModel, QuotaItemEditView>();
           
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            
            var mainViewModel = new MainViewModel();            
            displayRootRegistry.ShowPresentation(mainViewModel);
            
        }

        protected override void OnExit(ExitEventArgs e)
        {            
            displayRootRegistry.UnregisterWindowType<MainViewModel>();
            displayRootRegistry.UnregisterWindowType<DetailsViewModel>();
            displayRootRegistry.UnregisterWindowType<ClientViewModel>();
            displayRootRegistry.UnregisterWindowType<DIC_ClientViewModel>();
            displayRootRegistry.UnregisterWindowType<DIC_QuotationViewModel>();
            displayRootRegistry.UnregisterWindowType<DIC_addItemViewModel>();
            displayRootRegistry.UnregisterWindowType<DIC_addDescriptionViewModel>();
            displayRootRegistry.UnregisterWindowType<DIC_ContractorViewModel>();
            displayRootRegistry.UnregisterWindowType<DIC_SupplierViewModel>();
            displayRootRegistry.UnregisterWindowType<DIC_DepthViewModel>();
            displayRootRegistry.UnregisterWindowType<QuotationViewModel>();
            displayRootRegistry.UnregisterWindowType<SelectClientViewModel>();
            displayRootRegistry.UnregisterWindowType<SelectQuotaViewModel>();
            displayRootRegistry.UnregisterWindowType<QuotaOtherViewModel>();
            displayRootRegistry.UnregisterWindowType<PaymentViewModel>();
            displayRootRegistry.UnregisterWindowType<InvoiceViewModel>();
            displayRootRegistry.UnregisterWindowType<WorkOrderViewModel>();
            displayRootRegistry.UnregisterWindowType<MaterialProfitViewModel>();
            displayRootRegistry.UnregisterWindowType<LabourProfitViewModel>();
            displayRootRegistry.UnregisterWindowType<FotoViewModel>();
            displayRootRegistry.UnregisterWindowType<ExpensesViewModel>();
            displayRootRegistry.UnregisterWindowType<DebtsViewModel>();
            displayRootRegistry.UnregisterWindowType<DebtsPaymentViewModel>();
            displayRootRegistry.UnregisterWindowType<MessageViewModel>();
            displayRootRegistry.UnregisterWindowType<DeliveryViewModel>();
            displayRootRegistry.UnregisterWindowType<GeneratedViewModel>();
            displayRootRegistry.UnregisterWindowType<QuotaItemEditViewModel>();
            
            base.OnExit(e);
        }
    }   

}
