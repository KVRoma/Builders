using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class Generated
    {
        public int Id { get; set; }
        public int ClientId { get; set; }

        public string Color { get; set; } = "Black";
        public bool IsEnable { get; set; } = true;

        public ICollection<GeneratedMaterial> GeneratedMaterials  { get; set; }
        public ICollection<GeneratedAccessories> GeneratedAccessories { get; set; }
        public ICollection<GeneratedStairs> GeneratedStairs { get; set; }
        public ICollection<GeneratedSuplementary> GeneratedSuplementaries { get; set; }
        public ICollection<GeneratedMolding> GeneratedMoldings { get; set; }
        public ICollection<GeneratedFlood> GeneratedFloods { get; set; }
        public ICollection<GeneratedList> GeneratedLists { get; set; }
        public Generated()
        {
            GeneratedMaterials = new List<GeneratedMaterial>();
            GeneratedAccessories = new List<GeneratedAccessories>();
            GeneratedStairs = new List<GeneratedStairs>();
            GeneratedSuplementaries = new List<GeneratedSuplementary>();
            GeneratedMoldings = new List<GeneratedMolding>();
            GeneratedFloods = new List<GeneratedFlood>();
            GeneratedLists = new List<GeneratedList>();
        }
    }
}
