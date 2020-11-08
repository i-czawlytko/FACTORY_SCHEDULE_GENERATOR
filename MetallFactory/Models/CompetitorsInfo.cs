using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MetallFactory.Models
{
    public class CompetitorsInfo
    {
        public int MatId { get; set; }
        public Dictionary<int,int> MachinesOps { get; set; }
    }
}
