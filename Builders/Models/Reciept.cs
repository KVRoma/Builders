using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class Reciept
    {
        public int? Id { get; set; }
        public int? QuotaId { get; set; }
        public string NumberQuota { get; set; }
        public string Number 
        {
            get { return "R" + (Id + 1000).ToString() + NumberQuota; }
        }
        public int? PayNumber { get; set; }
    }
}
