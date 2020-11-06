﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MetallFactory.Models
{
    public class ScheduleRow
    {
        public int PartyId { get; set; }
        public int MaterialId { get; set; }
        public int MachineId { get; set; }
        public int StartTime { get; set; }
        public int EndTime { get; set; }
    }
}
