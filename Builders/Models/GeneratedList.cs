using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class GeneratedList
    {
        public int? Id { get; set; }
        public string Groupe { get; set; }
        public string GradeLevel { get; set; }
        public string Partition { get; set; }
        public string Aditional { get; set; }
        public string Name { get; set; }
        public decimal Count { get; set; }
        public string FullRoom 
        {
            get { return GradeLevel + " " + Partition + " " + Aditional; }
        }

        public int? GeneratedId { get; set; }
        public Generated Generated { get; set; }
    }
}
