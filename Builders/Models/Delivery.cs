using System;
using System.Collections.Generic;
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
        public int? SupplierId { get; set; }
        public string SupplierName { get; set; }
        public string OrderNumber { get; set; }

        public bool IsEnabled { get; set; } = false;
        public string Color { get; set; } = "Black";

        public ICollection<DeliveryMaterial> DeliveryMaterial { get; set; }
       
        public Delivery()
        {
            DeliveryMaterial = new List<DeliveryMaterial>();
            
        }


    }
}
