using MetallFactory.ViewModels;
using Microsoft.AspNetCore.Hosting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;


namespace MetallFactory.Models
{
    public class ScheduleGenerator
    {
        private IRepository repository;

        private readonly IWebHostEnvironment _webHostEnvironment;
        private List<List<ScheduleRow>> schedules_all;

        public ScheduleGenerator(IRepository repo, IWebHostEnvironment webHostEnvironment)
        {
            repository = repo;
            _webHostEnvironment = webHostEnvironment;
            schedules_all = new List<List<ScheduleRow>>();
        }
        public List<ScheduleRow> Generate(List<TIStructured> time_data )
        {
            repository.Load();
            var _schedule = new List<ScheduleRow>();
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
                    var current_machine_info = time_data.FirstOrDefault(c => c.MachineId == m);
                    bool party_was_found = false;
                    foreach(var e in current_machine_info?.TimeDict ?? new List<(int, int)>())
                    {
                        var party = parties.FirstOrDefault(x => x.MaterialId == e.Item2);
                        if (party != null)
                        {
                            if (parties.Remove(party))
                            {
                                _schedule.Add(new ScheduleRow
                                {
                                    PartyId = party.Id,
                                    MaterialId = e.Item2,
                                    MachineId = current_machine_info.MachineId,
                                    StartTime = current_time,
                                    EndTime = (current_time + e.Item1)
                                });

                                next_loading[current_machine_info.MachineId] += e.Item1;
                                party_was_found = true;
                                break;
                            }
                        }                     
                    }
                    if(!party_was_found) next_loading.Remove(m);
                }
            }
            return _schedule;

        }
        public void GenerateAll()
        {
            var combos = repository.AllCombinations;
            for(int i = 0; i < combos.Count; i++)
            {
                var combo = combos[i];
                this.schedules_all.Add(this.Generate(combo));
            }
            
        }

        public IEnumerable<ScheduleVM> GetAllSchedulesVM()
        {
            return schedules_all.Select(s => new ScheduleVM { TotalTime = s.Max(x => x.EndTime) }).OrderBy(s => s.TotalTime);
        }

        public List<List<ScheduleRow>> GetAllSchedules()
        {
            return schedules_all.OrderBy(s => s.Max(x=>x.EndTime)).ToList();
        }

        public IEnumerable<ScheduleRowVM> GetAnySchedule(List<ScheduleRow> _schedule)
        {
            var result = from sr in _schedule
                         join m in repository.Machines on sr.MachineId equals m.Id
                         join mat in repository.Materials on sr.MaterialId equals mat.Id
                         select new ScheduleRowVM { PartyId = sr.PartyId, MachineName = m.Name, MaterialName = mat.Name, StartTime = sr.StartTime, EndTime = sr.EndTime };
            return result;
        }
        public void ExportToXlxs(int idx)
        {
            string contentRootPath = _webHostEnvironment.ContentRootPath;

            string path = Path.Combine(contentRootPath, "output", "schedule.xlsx");

            FileInfo f = new FileInfo(path);
            if (f.Exists) f.Delete();
            using (ExcelPackage ep = new ExcelPackage(f))
            {
                ExcelWorksheet sch = ep.Workbook.Worksheets.Add("Schedule");
                sch.Cells[1, 1].Value = "ID партии"; ;
                sch.Cells[1, 2].Value = "Материал";
                sch.Cells[1, 3].Value = "Машина";
                sch.Cells[1, 4].Value = "Начало";
                sch.Cells[1, 5].Value = "Окончание";

                this.GenerateAll();
                var schedule = this.GetAllSchedules()[idx];
                var src = this.GetAnySchedule(schedule);

                int counter = 2;
                foreach (var e in src)
                {
                    sch.Cells[counter, 1].Value = e.PartyId;
                    sch.Cells[counter, 2].Value = e.MaterialName;
                    sch.Cells[counter, 3].Value = e.MachineName;
                    sch.Cells[counter, 4].Value = e.StartTime;
                    sch.Cells[counter, 5].Value = e.EndTime;
                    counter++;
                }
                ep.Save();
            }
        }
    }
}
