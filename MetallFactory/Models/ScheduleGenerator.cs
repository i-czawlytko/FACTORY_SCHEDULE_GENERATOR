﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;

namespace MetallFactory.Models
{
    public class ScheduleGenerator
    {
        private IRepository repository;

        private List<ScheduleRow> schedule;

        public ScheduleGenerator(IRepository repo)
        {
            repository = repo;
            repository.Load();
            schedule = new List<ScheduleRow>();
        }
        public List<ScheduleRow> Generate()
        {
            var parties = repository.Parties;
            int current_time;

            Dictionary<int, int> next_loading = new Dictionary<int, int>();
            foreach(var m in repository.Machines)
            {
                next_loading.Add(m.Id,0);
            }

            while (parties.Any())
            {
                current_time = next_loading.Select(x=>x.Value).Min();
                var free_machines = next_loading.Where(x=> x.Value == current_time).Select(a => a.Key).ToList();
                foreach (var m in free_machines)
                {
                    var current_machine_info = repository.StructuredTimes.FirstOrDefault(c => c.MachineId == m);
                    bool party_was_found = false;
                    foreach(var e in current_machine_info.TimeDict)
                    {
                        var party = parties.FirstOrDefault(x => x.MaterialId == e.Value);
                        if (party != null)
                        {
                            if (parties.Remove(party))
                            {
                                schedule.Add(new ScheduleRow {
                                    Party=e.Value,
                                    Machine=current_machine_info.MachineId,
                                    StartTime=current_time,
                                    EndTime=(current_time+e.Key)});

                                next_loading[current_machine_info.MachineId] += e.Key;
                                party_was_found = true;
                                break;
                            }
                        }                        
                    }
                    if(!party_was_found) next_loading.Remove(m);
                }
            }
            return schedule;

        }
    }
}
