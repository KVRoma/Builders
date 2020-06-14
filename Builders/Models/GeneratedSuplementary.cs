using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class GeneratedSuplementary
    {
        public int? Id { get; set; }
        public decimal FurnitureHandelingRoom { get; set; }
        public decimal RateDisposal { get; set; }
        public string DeliveryName { get; set; }
        public decimal DeliveryQty { get; set; }
        public string GeneralNotes { get; set; }

        public int? GeneratedId { get; set; }
        public Generated Generated { get; set; }
    }
}
