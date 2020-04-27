using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class DIC_DescriptionQuotation
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public decimal Price { get; set; }
        public decimal Cost { get; set; }

        public int? ItemId { get; set; }
        public DIC_ItemQuotation DIC_ItemQuotation  { get; set; }
    }
}
