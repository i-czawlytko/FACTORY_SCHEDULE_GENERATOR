using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MetallFactory.ViewModels
{
    public class ScheduleRowVM
    {
        public int PartyId { get; set; }
        public string MachineName { get; set; }
        public string MaterialName { get; set; }
        public int StartTime { get; set; }
        public int EndTime { get; set; }
    }
}
