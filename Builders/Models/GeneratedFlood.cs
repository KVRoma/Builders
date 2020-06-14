using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class GeneratedFlood
    {
        public int? Id { get; set; }
        public string RoomDescription { get; set; }
        public decimal Size { get; set; }
        public string Depth { get; set; }
        public decimal QtyNovoplan { get; set; }
        public string Concatenar { get; set; }

        public int? GeneratedId { get; set; }
        public Generated Generated { get; set; }
    }
}
