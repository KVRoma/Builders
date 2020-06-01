using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class Expenses
    {
        public int Id { get; set; }
        public string Type { get; set; }
        public DateTime Date { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public decimal Amounts { get; set; }

        public DateTime DatePaid { get; set; } = DateTime.MinValue;
        public string NotesPaid { get; set; }
        public decimal AmountPaid { get; set; } = 0m;

        public bool Payment { get; set; } = false;
        public string Color { get; set; } = "Red";

        public string CompanyName { get; set; } = "CMO";
    }
}
