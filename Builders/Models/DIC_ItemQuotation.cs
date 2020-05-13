using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class DIC_ItemQuotation
    {
        public int? Id { get; set; }
        public string Name { get; set; }

        public int? SupplierId { get; set; }
        public string Color { get; set; } = "Black";
        
        public int? GroupeId { get; set; }
        public DIC_GroupeQuotation DIC_GroupeQuotation  { get; set; }

        public ICollection<DIC_DescriptionQuotation> DIC_DescriptionQuotation  { get; set; }
        public DIC_ItemQuotation()
        {
            DIC_DescriptionQuotation = new List<DIC_DescriptionQuotation>();
        }

    }
}
