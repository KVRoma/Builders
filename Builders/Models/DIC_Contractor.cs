using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class DIC_Contractor
    {
        public int? Id { get; set; }
        public string NameAndPhone { get; set; }
        public string Name { get; set; }
        public string Specialty { get; set; }
        public string SpecName 
        { 
            get { return Name + ((Specialty != "") ? (" (") : ("")) + Specialty + ((Specialty != "") ? (") ") : ("")); } 
        }
        [Required]
        public string Color { get; set; } = "Black";
    }
}
