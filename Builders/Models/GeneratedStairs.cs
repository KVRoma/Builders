using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class GeneratedStairs
    {
        public int? Id { get; set; }
        public string GradeLevel { get; set; }
        public string TypeStairs { get; set; }
        public decimal QtyStairs { get; set; }
        public decimal LenghtStairs { get; set; }
        public decimal QtyStairsLenght { get; set; }
        public string TypeLeveling { get; set; }
        public decimal QtyLeveling { get; set; }
        public string Demolition { get; set; }

        public int? GeneratedId { get; set; }
        public Generated Generated { get; set; }
    }
}
