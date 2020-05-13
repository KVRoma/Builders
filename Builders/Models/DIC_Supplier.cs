using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class DIC_Supplier
    {
        public int Id { get; set; }
        public string Supplier { get; set; }
        public string Brands { get; set; }
        public string Rep { get; set; }
        public string RepLine { get; set; }
        public string OrderLine { get; set; }
        public string Address { get; set; }
        public string Page { get; set; }
        public string Email { get; set; }
        public string Hours { get; set; }

        public int? GroupeId { get; set; }
        public DIC_GroupeQuotation DIC_GroupeQuotation { get; set; }
    }
}
