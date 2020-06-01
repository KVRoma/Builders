using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class Delivery
    {
        public int Id { get; set; }
        public DateTime DateCreating { get; set; }

        public int? QuotaId { get; set; }
        public string NumberQuota { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string PhoneNumber { get; set; }
        public string Email { get; set; }
        public string FullName
        {
            get { return FirstName + " " + LastName; }
        }
        public string FullSearch
        {
            get { return NumberQuota + " " + FirstName + " " + LastName + " " + PhoneNumber + " " + Email; }
        }
       
        public string NameComboBox
        {
            get { return NumberQuota + " " + FullName; }
        }
        public int? SupplierId { get; set; }
        public string SupplierName { get; set; }
        public string OrderNumber { get; set; }
        public decimal AmountDelivery { get; set; } = 0m;

        public bool IsArchive { get; set; } = false;
        public bool IsEnabled { get; set; } = false;
        public string Color { get; set; } = "Red";

        public string CompanyName { get; set; } = "CMO";

        public ICollection<DeliveryMaterial> DeliveryMaterial { get; set; }
       
        public Delivery()
        {
            DeliveryMaterial = new List<DeliveryMaterial>();
            
        }


    }
}
