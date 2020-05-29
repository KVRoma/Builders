using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class WorkOrder_Work
    {
        public int? Id { get; set; }
        public string Area { get; set; }
        public string Room { get; set; }
        public string Existing { get; set; }
        public string NewFloor { get; set; }
        public string Furniture { get; set; }
        public string Misc { get; set; }

        public string Contractor { get; set; }
        public string Color { get; set; }

        public int? WorkOrderId { get; set; }
        public virtual WorkOrder WorkOrder { get; set; }
    }
}
