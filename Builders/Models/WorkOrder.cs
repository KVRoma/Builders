using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class WorkOrder
    {
        public int? Id { get; set; }
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
        public DateTime DateWork { get; set; }
        public DateTime DateServices { get; set; }
        public DateTime DateCompletion { get; set; }
        public string Parking { get; set; }
        public string Trim { get; set; }
        public string Colour { get; set; }
        public string LF { get; set; }
        public string Baseboard { get; set; }
        public string ReplacingYesNo { get; set; }
        public int? ReplacingQuantity { get; set; }

        public string CompanyName { get; set; } = "CMO";

        public virtual ICollection<WorkOrder_Work> WorkOrder_Works { get; set; }
        public virtual ICollection<WorkOrder_Installation> WorkOrder_Installations { get; set; }
        public virtual ICollection<WorkOrder_Accessories> WorkOrder_Accessories { get; set; }
        public virtual ICollection<WorkOrder_Contractor> WorkOrder_Contractors { get; set; }

        public WorkOrder()
        {
            WorkOrder_Works = new List<WorkOrder_Work>();
            WorkOrder_Accessories = new List<WorkOrder_Accessories>();
            WorkOrder_Installations = new List<WorkOrder_Installation>();
            WorkOrder_Contractors = new List<WorkOrder_Contractor>();
        }
    }
}
