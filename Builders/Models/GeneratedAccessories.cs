using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class GeneratedAccessories
    {
        public int? Id { get; set; }
        public string NewFloor { get; set; }
        public decimal TotalFloor { get; set; }
        public string AccessoriesFloor { get; set; }
        public string TypeAccessoriesFloor { get; set; }
        public decimal TotalAccessoriesFloor { get; set; }

        public int? GeneratedId { get; set; }
        public Generated Generated { get; set; }
    }
}
