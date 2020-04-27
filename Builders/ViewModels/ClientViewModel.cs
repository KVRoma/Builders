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
    public class ClientViewModel : ViewModel
    {
        public string WindowName { get; } = "Builder - (Client)";
        private BuilderContext db;

        private int id;
        private DateTime dateRegistration;
        private DIC_TypeOfClient typeOfClientSelect;
        private IEnumerable<DIC_TypeOfClient> typeOfClients;
        private string companyName;
        private string primaryFirstName;
        private string primaryLastName;
        private string primaryPhoneNumber;
        private string primaryEmail;
        private string secondaryFirstName;
        private string secondaryLastName;
        private string secondaryPhoneNumber;
        private string secondaryEmail;
        private string addressBillStreet;
        private string addressBillCity;
        private string addressBillProvince;
        private string addressBillPostalCode;
        private string addressBillCountry;
        private string addressSiteStreet;
        private string addressSiteCity;
        private string addressSiteProvince;
        private string addressSitePostalCode;
        private string addressSiteCountry;
        private DIC_HearAboutsUs hearAboutsUsSelect;
        private IEnumerable<DIC_HearAboutsUs> hearAboutsUse;
        private string specify;
        private string notes;
        private EnumClient result;

        public int Id
        {
            get { return id; }
            set
            {
                id = value;
                OnPropertyChanged(nameof(Id));
            }
        }
        public DateTime DateRegistration
        {
            get { return dateRegistration; }
            set
            {
                dateRegistration = value;
                OnPropertyChanged(nameof(DateRegistration));
            }
        }
        public DIC_TypeOfClient TypeOfClientSelect
        {
            get { return typeOfClientSelect; }
            set
            {
                typeOfClientSelect = value;
                OnPropertyChanged(nameof(TypeOfClientSelect));
            }
        }
        public IEnumerable<DIC_TypeOfClient> TypeOfClients
        {
            get { return typeOfClients; }
            set
            {
                typeOfClients = value;
                OnPropertyChanged(nameof(TypeOfClients));
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
        public string PrimaryFirstName
        {
            get { return primaryFirstName; }
            set
            {
                primaryFirstName = value;
                OnPropertyChanged(nameof(PrimaryFirstName));
            }
        }
        public string PrimaryLastName
        {
            get { return primaryLastName; }
            set
            {
                primaryLastName = value;
                OnPropertyChanged(nameof(PrimaryLastName));
            }
        }
        public string PrimaryPhoneNumber
        {
            get { return primaryPhoneNumber; }
            set
            {
                primaryPhoneNumber = value;
                OnPropertyChanged(nameof(PrimaryPhoneNumber));
            }
        }
        public string PrimaryEmail
        {
            get { return primaryEmail; }
            set
            {
                primaryEmail = value;
                OnPropertyChanged(nameof(PrimaryEmail));
            }
        }
        public string SecondaryFirstName
        {
            get { return secondaryFirstName; }
            set
            {
                secondaryFirstName = value;
                OnPropertyChanged(nameof(SecondaryFirstName));
            }
        }
        public string SecondaryLastName
        {
            get { return secondaryLastName; }
            set
            {
                secondaryLastName = value;
                OnPropertyChanged(nameof(SecondaryLastName));
            }
        }
        public string SecondaryPhoneNumber
        {
            get { return secondaryPhoneNumber; }
            set
            {
                secondaryPhoneNumber = value;
                OnPropertyChanged(nameof(SecondaryPhoneNumber));
            }
        }
        public string SecondaryEmail
        {
            get { return secondaryEmail; }
            set
            {
                secondaryEmail = value;
                OnPropertyChanged(nameof(SecondaryEmail));
            }
        }
        public string AddressBillStreet
        {
            get { return addressBillStreet; }
            set
            {
                addressBillStreet = value;
                OnPropertyChanged(nameof(AddressBillStreet));
            }
        }
        public string AddressBillCity
        {
            get { return addressBillCity; }
            set
            {
                addressBillCity = value;
                OnPropertyChanged(nameof(AddressBillCity));
            }
        }
        public string AddressBillProvince
        {
            get { return addressBillProvince; }
            set
            {
                addressBillProvince = value;
                OnPropertyChanged(nameof(AddressBillProvince));
            }
        }
        public string AddressBillPostalCode
        {
            get { return addressBillPostalCode; }
            set
            {
                addressBillPostalCode = value;
                OnPropertyChanged(nameof(AddressBillPostalCode));
            }
        }
        public string AddressBillCountry
        {
            get { return addressBillCountry; }
            set
            {
                addressBillCountry = value;
                OnPropertyChanged(nameof(AddressBillCountry));
            }
        }
        public string AddressSiteStreet
        {
            get { return addressSiteStreet; }
            set
            {
                addressSiteStreet = value;
                OnPropertyChanged(nameof(AddressSiteStreet));
            }
        }
        public string AddressSiteCity
        {
            get { return addressSiteCity; }
            set
            {
                addressSiteCity = value;
                OnPropertyChanged(nameof(AddressSiteCity));
            }
        }
        public string AddressSiteProvince
        {
            get { return addressSiteProvince; }
            set
            {
                addressSiteProvince = value;
                OnPropertyChanged(nameof(AddressSiteProvince));
            }
        }
        public string AddressSitePostalCode
        {
            get { return addressSitePostalCode; }
            set
            {
                addressSitePostalCode = value;
                OnPropertyChanged(nameof(AddressSitePostalCode));
            }
        }
        public string AddressSiteCountry
        {
            get { return addressSiteCountry; }
            set
            {
                addressSiteCountry = value;
                OnPropertyChanged(nameof(AddressSiteCountry));
            }
        }
        public DIC_HearAboutsUs HearAboutsUsSelect
        {
            get { return hearAboutsUsSelect; }
            set
            {
                hearAboutsUsSelect = value;
                OnPropertyChanged(nameof(HearAboutsUsSelect));
            }
        }
        public IEnumerable<DIC_HearAboutsUs> HearAboutsUse
        {
            get { return hearAboutsUse; }
            set
            {
                hearAboutsUse = value;
                OnPropertyChanged(nameof(HearAboutsUse));
            }
        }
        public string Specify
        {
            get { return specify; }
            set
            {
                specify = value;
                OnPropertyChanged(nameof(Specify));
            }
        }
        public string Notes
        {
            get { return notes; }
            set
            {
                notes = value;
                OnPropertyChanged(Notes);
            }
        }
        public EnumClient Result
        {
            get { return result; }
            set
            {
                result = value;
                OnPropertyChanged(nameof(Result));
            }
        }

        private Command _okCommand;

        public Command OkCommand => _okCommand ?? (_okCommand = new Command(obj =>
        {
            if (DateRegistration != null)
            {
                switch (Result)
                {
                    case EnumClient.Add:
                        {
                            Client client = new Client()
                            {
                                DateRegistration = DateRegistration,
                                TypeOfClient = TypeOfClientSelect?.Name ?? "",
                                CompanyName = CompanyName,
                                PrimaryFirstName = PrimaryFirstName ?? "",
                                PrimaryLastName = PrimaryLastName ?? "",
                                PrimaryPhoneNumber = PrimaryPhoneNumber ?? "",
                                PrimaryEmail = PrimaryEmail ?? "",
                                SecondaryFirstName = SecondaryFirstName,
                                SecondaryLastName = SecondaryLastName,
                                SecondaryPhoneNumber = SecondaryPhoneNumber,
                                SecondaryEmail = SecondaryEmail,
                                AddressBillStreet = AddressBillStreet,
                                AddressBillCity = AddressBillCity,
                                AddressBillProvince = AddressBillProvince,
                                AddressBillPostalCode = AddressBillPostalCode,
                                AddressBillCountry = AddressBillCountry,
                                AddressSiteStreet = AddressSiteStreet,
                                AddressSiteCity = AddressSiteCity,
                                AddressSiteProvince = AddressSiteProvince,
                                AddressSitePostalCode = AddressSitePostalCode,
                                AddressSiteCountry = AddressSiteCountry,
                                HearAboutUs = HearAboutsUsSelect?.Name ?? "",
                                Specify = Specify,
                                Notes = Notes
                            };
                            db.Clients.Add(client);
                            db.SaveChanges();
                            Result = EnumClient.Null;
                        }
                        break;
                    case EnumClient.Ins:
                        {
                            var client = db.Clients.Find(Id);
                            client.DateRegistration = DateRegistration;
                            client.TypeOfClient = TypeOfClientSelect?.Name ?? "";
                            client.CompanyName = CompanyName;
                            client.PrimaryFirstName = PrimaryFirstName ?? "";
                            client.PrimaryLastName = PrimaryLastName ?? "";
                            client.PrimaryPhoneNumber = PrimaryPhoneNumber ?? "";
                            client.PrimaryEmail = PrimaryEmail ?? "";
                            client.SecondaryFirstName = SecondaryFirstName;
                            client.SecondaryLastName = SecondaryLastName;
                            client.SecondaryPhoneNumber = SecondaryPhoneNumber;
                            client.SecondaryEmail = SecondaryEmail;
                            client.AddressBillStreet = AddressBillStreet;
                            client.AddressBillCity = AddressBillCity;
                            client.AddressBillProvince = AddressBillProvince;
                            client.AddressBillPostalCode = AddressBillPostalCode;
                            client.AddressBillCountry = AddressBillCountry;
                            client.AddressSiteStreet = AddressSiteStreet;
                            client.AddressSiteCity = AddressSiteCity;
                            client.AddressSiteProvince = AddressSiteProvince;
                            client.AddressSitePostalCode = AddressSitePostalCode;
                            client.AddressSiteCountry = AddressSiteCountry;
                            client.HearAboutUs = HearAboutsUsSelect?.Name ?? "";
                            client.Specify = Specify;
                            client.Notes = Notes;

                            db.Entry(client).State = EntityState.Modified;
                            db.SaveChanges();
                            Result = EnumClient.Null;
                        }
                        break;
                    case EnumClient.Del:
                        {
                        }
                        break;
                    case EnumClient.Null:
                        break;
                }
                
                if (obj is System.Windows.Window)
                {
                    (obj as System.Windows.Window).Close();
                }
            }            
        }));

        public ClientViewModel(ref BuilderContext context, EnumClient result)
        {
            db = context;
            Result = result;
            db.DIC_TypeOfClients.Load();
            db.DIC_HearAboutsUse.Load();
            TypeOfClients = db.DIC_TypeOfClients.Local.ToBindingList();
            HearAboutsUse = db.DIC_HearAboutsUse.Local.ToBindingList();
        }
    }
}
