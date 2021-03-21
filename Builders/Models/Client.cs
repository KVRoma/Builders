using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class Client
    {
        public int Id { get; set; }
        public DateTime DateRegistration { get; set; }
        public string NumberClient
        {
            get { return "C" + (Id + 1000).ToString(); }
        }
        public string NumberComboBox
        {
            get { return NumberClient + " " + PrimaryFullName; }
        }
        public string TypeOfClient { get; set; }
        public string CompanyName { get; set; }
        public string PrimaryFirstName { get; set; }
        public string PrimaryLastName { get; set; }
        public string PrimaryPhoneNumber { get; set; }
        public string PrimaryEmail { get; set; }
        public string PrimaryFullName
        {
            get { return PrimaryFirstName + " " + PrimaryLastName; }
        }
        public string FullSearch
        {
            get { return TypeOfClient + " " + CompanyName + " " + PrimaryFirstName + " " + PrimaryLastName + " " + PrimaryPhoneNumber + " " + PrimaryEmail; }
        }
        public string SecondaryFirstName { get; set; }
        public string SecondaryLastName { get; set; }
        public string SecondaryPhoneNumber { get; set; }
        public string SecondaryEmail { get; set; }
        public string AddressBillStreet { get; set; }
        public string AddressBillCity { get; set; }
        public string AddressBillProvince { get; set; }
        public string AddressBillPostalCode { get; set; }
        public string AddressBillCountry { get; set; }
        public string AddressBillFull
        {
            get
            {
                return AddressBillStreet + ", " +
                       AddressBillCity + ", " +
                       AddressBillProvince + ", " +
                       AddressBillPostalCode + ", " +
                       AddressBillCountry;
            }
        }
        public string AddressSiteStreet { get; set; }
        public string AddressSiteCity { get; set; }
        public string AddressSiteProvince { get; set; }
        public string AddressSitePostalCode { get; set; }
        public string AddressSiteCountry { get; set; }
        public string AddressSiteFull
        {
            get
            {
                return AddressSiteStreet + ", " +
                       AddressSiteCity + ", " +
                       AddressSiteProvince + ", " +
                       AddressSitePostalCode + ", " +
                       AddressSiteCountry;
            }
        }
        public string HearAboutUs { get; set; }
        public string Specify { get; set; }
        public string Notes { get; set; }
        
    }
}
