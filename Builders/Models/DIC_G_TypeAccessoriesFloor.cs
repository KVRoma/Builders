using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class DIC_G_TypeAccessoriesFloor
    {
        public int? Id { get; set; }
        public string Name { get; set; }

        public int? DIC_G_AccessoriesFloorId { get; set; }
        public DIC_G_AccessoriesFloor DIC_G_AccessoriesFloor { get; set; }
    }
}
