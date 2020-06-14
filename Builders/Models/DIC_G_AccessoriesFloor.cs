using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Builders.Models
{
    public class DIC_G_AccessoriesFloor
    {
        public int? Id { get; set; }
        public string Name { get; set; }

        public ICollection<DIC_G_TypeAccessoriesFloor> DIC_G_TypeAccessoriesFloors { get; set; }
        public DIC_G_AccessoriesFloor() 
        {
            DIC_G_TypeAccessoriesFloors = new List<DIC_G_TypeAccessoriesFloor>();
        }
            
    }
}
