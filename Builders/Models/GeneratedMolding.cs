using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class GeneratedMolding
    {
        public int? Id { get; set; }
        public string MoldingName { get; set; }
        public string TypeMolding { get; set; }
        public decimal QtyMolding { get; set; }
        public decimal HeightMolding { get; set; }
        public decimal BaseboardMaterial { get; set; }
        public string Painting { get; set; }

        public int? GeneratedId { get; set; }
        public Generated Generated { get; set; }
    }
}
