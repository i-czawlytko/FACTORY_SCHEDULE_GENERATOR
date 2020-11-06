using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MetallFactory.Models
{
    public class ScheduleRow
    {
        public int Party { get; set; }
        public int Machine { get; set; }
        public int StartTime { get; set; }
        public int EndTime { get; set; }
    }
}
