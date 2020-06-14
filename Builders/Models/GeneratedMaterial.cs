using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class GeneratedMaterial
    {
        public int? Id { get; set; }
        public string GradeLevel { get; set; }
        public string Partition { get; set; }
        public string Aditional { get; set; }
        public decimal TransitionR { get; set; }
        public decimal TransitionT { get; set; }
        public decimal TransitionOther { get; set; }
        public string NoteTransitionOther { get; set; }
        public string ExistingFloor { get; set; }
        public string NeedDemolition { get; set; }
        public string NewFloor { get; set; }
        public decimal LenghtFloor { get; set; }
        public decimal WidthFloor { get; set; }
        public decimal TotalFloor { get; set; }
        public decimal TotalFlooringMaterial { get; set; }
        public string TypeTrim { get; set; }
        public decimal QtyTrim { get; set; }

        public int? GeneratedId { get; set; }
        public Generated Generated { get; set; }
    }
}
