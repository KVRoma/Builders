using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public interface IdicQuotation
    {
        int Id { get; set; }
        string NameItem { get; set; }
        string NameDescription { get; set; }
        decimal Price { get; set; }
        decimal Cost { get; set; }
    }
}
