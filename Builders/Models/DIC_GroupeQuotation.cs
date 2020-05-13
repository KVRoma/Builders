using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class DIC_GroupeQuotation
    {
        public int? Id { get; set; }
        public string NameGroupe { get; set; }

        public ICollection<DIC_ItemQuotation> DIC_ItemQuotation { get; set; }
        public ICollection<DIC_Supplier> DIC_Supplier { get; set; }
        public DIC_GroupeQuotation()
        {
            DIC_ItemQuotation = new List<DIC_ItemQuotation>();
            DIC_Supplier = new List<DIC_Supplier>();
        }

    }
}
