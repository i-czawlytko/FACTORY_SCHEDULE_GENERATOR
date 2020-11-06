using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MetallFactory.Models
{
    public class TIStructured
    {
        public int MachineId { get; set; }
        public SortedDictionary<int,int> TimeDict { get; set; }
    }
}
