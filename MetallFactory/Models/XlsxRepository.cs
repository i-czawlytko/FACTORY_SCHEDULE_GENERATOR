
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Server;
using Microsoft.AspNetCore.Hosting;

namespace MetallFactory.Models
{
    public class XlsxRepository: IRepository
    {
        public List<Material> Materials { get; set; }
        public List<Machine> Machines { get; set; }
        public List<Party> Parties { get; set; }
        public List<TimeInfo> Times { get; set; }
        public List<TIStructured> StructuredTimes { get; set; }
        public List<CompetitorsInfo> Competitors { get; set; }

        private readonly IWebHostEnvironment _webHostEnvironment;

        public XlsxRepository(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
            this.LoadMaterials();
            this.LoadMachines();
            this.LoadParties();
            this.LoadTimes();
            this.TIRestructure();
            this.LoadCompetitors();
        }

        private void LoadMaterials()
        {
            Materials = new List<Material>();

            string contentRootPath = _webHostEnvironment.ContentRootPath;

            string path = Path.Combine(contentRootPath, "xlsx_data", "nomenclatures.xlsx");

            byte[] bin = File.ReadAllBytes(path);

            using (MemoryStream stream = new MemoryStream(bin))
            using (OfficeOpenXml.ExcelPackage excelPackage = new ExcelPackage(stream))
            {
                foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                {
                    for (int i = worksheet.Dimension.Start.Row+1; i <= worksheet.Dimension.End.Row; i++)
                    {
                        Material mat = new Material();

                        mat.Id = int.Parse(worksheet.Cells[i, 1].Value.ToString());
                        mat.Name = worksheet.Cells[i, 2].Value.ToString();
                        Materials.Add(mat);
                    }
                }
            }
        }
        
        private void LoadMachines()
        {
            Machines = new List<Machine>();

            string contentRootPath = _webHostEnvironment.ContentRootPath;

            string path = Path.Combine(contentRootPath, "xlsx_data", "machine_tools.xlsx");

            byte[] bin = File.ReadAllBytes(path);

            using (MemoryStream stream = new MemoryStream(bin))
            using (OfficeOpenXml.ExcelPackage excelPackage = new ExcelPackage(stream))
            {
                foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                {
                    for (int i = worksheet.Dimension.Start.Row+1; i <= worksheet.Dimension.End.Row; i++)
                    {
                        Machine machine = new Machine();

                        machine.Id = int.Parse(worksheet.Cells[i, 1].Value.ToString());
                        machine.Name = worksheet.Cells[i, 2].Value.ToString();
                        Machines.Add(machine);
                    }
                }
            }
        }

        private void LoadParties()
        {
            Parties = new List<Party>();

            string contentRootPath = _webHostEnvironment.ContentRootPath;

            string path = Path.Combine(contentRootPath, "xlsx_data", "parties.xlsx");

            byte[] bin = File.ReadAllBytes(path);

            using (MemoryStream stream = new MemoryStream(bin))
            using (OfficeOpenXml.ExcelPackage excelPackage = new ExcelPackage(stream))
            {
                foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                {
                    for (int i = worksheet.Dimension.Start.Row + 1; i <= worksheet.Dimension.End.Row; i++)
                    {
                        Party party = new Party();

                        party.Id = int.Parse(worksheet.Cells[i, 1].Value.ToString());
                        party.MaterialId = int.Parse(worksheet.Cells[i, 2].Value.ToString());
                        Parties.Add(party);
                    }
                }
            }
        }

        private void LoadTimes()
        {
            Times = new List<TimeInfo>();

            string contentRootPath = _webHostEnvironment.ContentRootPath;

            string path = Path.Combine(contentRootPath, "xlsx_data", "times.xlsx");

            byte[] bin = File.ReadAllBytes(path);

            using (MemoryStream stream = new MemoryStream(bin))
            using (OfficeOpenXml.ExcelPackage excelPackage = new ExcelPackage(stream))
            {
                foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                {
                    for (int i = worksheet.Dimension.Start.Row + 1; i <= worksheet.Dimension.End.Row; i++)
                    {
                        TimeInfo tinfo = new TimeInfo();

                        tinfo.MachineId = int.Parse(worksheet.Cells[i, 1].Value.ToString());
                        tinfo.MaterialId = int.Parse(worksheet.Cells[i, 2].Value.ToString());
                        tinfo.OperationTime = int.Parse(worksheet.Cells[i, 3].Value.ToString());
                        Times.Add(tinfo);
                    }
                }
            }
        }

        private void TIRestructure()
        {
            StructuredTimes = new List<TIStructured>();
            var machines = this.Times.Select(x => x.MachineId).Distinct();
            foreach(var m in machines)
            {
                TIStructured tis = new TIStructured();
                tis.MachineId = m;
                tis.TimeDict = new SortedDictionary<int, int>();

                var mats = this.Times.Where(t => t.MachineId == tis.MachineId);
                foreach (var mat in mats)
                {
                    tis.TimeDict.Add(mat.OperationTime,mat.MaterialId);
                }
                StructuredTimes.Add(tis);
            }
        }

        private void LoadCompetitors()
        {
            this.Competitors = new List<CompetitorsInfo>();
            var groups = from ti in this.Times
                      group ti by ti.MaterialId;
            foreach(var g in groups)
            {
                CompetitorsInfo ci = new CompetitorsInfo();
                ci.MachinesOps = new Dictionary<int, int>();

                ci.MatId = g.Key;
                foreach(var e in g)
                {
                    ci.MachinesOps.Add(e.MachineId,e.OperationTime);
                }
                this.Competitors.Add(ci);
            }
        }
        
    }
}
