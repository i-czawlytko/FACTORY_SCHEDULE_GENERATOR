
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Server;
using Microsoft.AspNetCore.Hosting;
using System.Text;
using System.Reflection;

namespace MetallFactory.Models
{
    public class XlsxRepository: IRepository
    {
        public List<Material> Materials { get; set; }
        public List<Machine> Machines { get; set; }
        public List<Party> Parties { get; set; }
        public List<TimeInfo> Times { get; set; }
        public List<TIStructured> StructuredTimes { get; set; }
        public Dictionary<int,Dictionary<int,int>> Competitors { get; set; }

        private StringBuilder errors;

        private readonly IWebHostEnvironment _webHostEnvironment;

        public XlsxRepository(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
            errors = new StringBuilder();
        }

        public void Load()
        {
            this.LoadMaterials();
            this.LoadMachines();
            this.LoadParties();
            this.LoadTimes();

            checkIDs();
            if (errors.Length > 0) throw new ExcelDataException(errors.ToString());

            this.TIRestructure();
            this.LoadCompetitors();

            
        }

        private void LoadMaterials()
        {
            Materials = new List<Material>();

            string contentRootPath = _webHostEnvironment.ContentRootPath;

            string filename = "nomenclatures.xlsx";

            string path = Path.Combine(contentRootPath, "xlsx_data", filename);

            byte[] bin = File.ReadAllBytes(path);

            using (MemoryStream stream = new MemoryStream(bin))
            using (OfficeOpenXml.ExcelPackage excelPackage = new ExcelPackage(stream))
            {
                foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                {
                    for (int i = worksheet.Dimension.Start.Row+1; i <= worksheet.Dimension.End.Row; i++)
                    {
                        Material mat = new Material();

                        if ( isCellNotEmptyAndNumber(worksheet, filename, i, 1) ) mat.Id = int.Parse( worksheet.Cells[i, 1].Value.ToString() );

                        if ( isCellNotEmpty(worksheet, filename, i, 2) ) mat.Name = worksheet.Cells[i, 2].Value.ToString();

                        Materials.Add(mat);
                    }
                }
            }
        }
        
        private void LoadMachines()
        {
            Machines = new List<Machine>();

            string contentRootPath = _webHostEnvironment.ContentRootPath;

            string filename = "machine_tools.xlsx";

            string path = Path.Combine(contentRootPath, "xlsx_data", filename);

            byte[] bin = File.ReadAllBytes(path);

            using (MemoryStream stream = new MemoryStream(bin))
            using (OfficeOpenXml.ExcelPackage excelPackage = new ExcelPackage(stream))
            {
                foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                {
                    for (int i = worksheet.Dimension.Start.Row+1; i <= worksheet.Dimension.End.Row; i++)
                    {
                        Machine machine = new Machine();

                        if (isCellNotEmptyAndNumber(worksheet, filename, i, 1)) machine.Id = int.Parse(worksheet.Cells[i, 1].Value.ToString());

                        if (isCellNotEmpty(worksheet, filename, i, 2)) machine.Name = worksheet.Cells[i, 2].Value.ToString();

                        Machines.Add(machine);
                    }
                }
            }
        }

        private void LoadParties()
        {
            Parties = new List<Party>();

            string contentRootPath = _webHostEnvironment.ContentRootPath;

            string filename = "parties.xlsx";

            string path = Path.Combine(contentRootPath, "xlsx_data", filename);

            byte[] bin = File.ReadAllBytes(path);

            using (MemoryStream stream = new MemoryStream(bin))
            using (OfficeOpenXml.ExcelPackage excelPackage = new ExcelPackage(stream))
            {
                foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                {
                    for (int i = worksheet.Dimension.Start.Row + 1; i <= worksheet.Dimension.End.Row; i++)
                    {
                        Party party = new Party();

                        if (isCellNotEmptyAndNumber(worksheet, filename, i, 1)) party.Id = int.Parse(worksheet.Cells[i, 1].Value.ToString());

                        if (isCellNotEmptyAndNumber(worksheet, filename, i, 2)) party.MaterialId = int.Parse(worksheet.Cells[i, 2].Value.ToString());

                        Parties.Add(party);
                    }
                }
            }
        }

        private void LoadTimes()
        {
            Times = new List<TimeInfo>();

            string contentRootPath = _webHostEnvironment.ContentRootPath;

            string filename = "times.xlsx";

            string path = Path.Combine(contentRootPath, "xlsx_data", filename);

            byte[] bin = File.ReadAllBytes(path);

            using (MemoryStream stream = new MemoryStream(bin))
            using (OfficeOpenXml.ExcelPackage excelPackage = new ExcelPackage(stream))
            {
                foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                {
                    for (int i = worksheet.Dimension.Start.Row + 1; i <= worksheet.Dimension.End.Row; i++)
                    {
                        TimeInfo tinfo = new TimeInfo();

                        if (isCellNotEmptyAndNumber(worksheet, filename, i, 1)) tinfo.MachineId = int.Parse(worksheet.Cells[i, 1].Value.ToString());

                        if (isCellNotEmptyAndNumber(worksheet, filename, i, 2)) tinfo.MaterialId = int.Parse(worksheet.Cells[i, 2].Value.ToString());

                        if (isCellNotEmptyAndNumber(worksheet, filename, i, 3)) tinfo.OperationTime = int.Parse(worksheet.Cells[i, 3].Value.ToString());

                        Times.Add(tinfo);
                    }
                }
            }
        }

        private bool isCellNotEmpty(ExcelWorksheet worksheet, string filename, int row, int col)
        {
            if (worksheet.Cells[row, col].Value != null)
            {
                return true;
            }
            else
            {
                errors.Append($"{filename}: строка {row}, столбец {col}. Пустая ячейка // ");
                return false;
            }
        }

        private bool isCellNotEmptyAndNumber(ExcelWorksheet worksheet, string filename, int row, int col)
        {
            if (worksheet.Cells[row, col].Value != null)
            {
                int num;
                if (int.TryParse(worksheet.Cells[row, col].Value.ToString(), out num))
                {
                    return true;
                }
                else
                {
                    errors.Append($"{filename}: строка {row}, столбец {col}. Не удалось преобразовать в число // ");
                    return false;
                }
            }
            else
            {
                errors.Append($"{filename}: строка {row}, столбец {col}. Пустая ячейка // ");
                return false;
            }
        }

        private void checkIDs()
        {
            if (this.Machines.Select(x => x.Id).GroupBy(v => v).Any(g => g.Count() > 1)) errors.Append("В machine_tools.xlsx имеются повторяющиеся значения идентификаторов // ");
            if (this.Materials.Select(x => x.Id).GroupBy(v => v).Any(g => g.Count() > 1)) errors.Append("В nomenclatures.xlsx имеются повторяющиеся значения идентификаторов // ");

            var parties_id = this.Parties.Select(x => x.Id);
            if (parties_id.GroupBy(v => v).Any(g => g.Count() > 1)) errors.Append("В parties.xlsx имеются повторяющиеся значения идентификаторов // ");

            var pairs_from_times = this.Times.Select(x => (x.MachineId, x.MaterialId));
            if (pairs_from_times.GroupBy(v => v).Any(g => g.Count() > 1)) errors.Append("В times.xlsx имеются повторяющиеся пары ID-материала и ID-оборудования // ");
        }

        private void TIRestructure()
        {
            StructuredTimes = new List<TIStructured>();

            var joined = from m in this.Machines
                         join ti in this.Times on m.Id equals ti.MachineId into j
                         from subti in j.DefaultIfEmpty()
                         select new { machine_id = m.Id, mat_id = subti?.MaterialId, op_time = subti?.OperationTime};

            var groups = from j in joined
                         group j by j.machine_id;

            foreach (var g in groups)
            {
                var tis = new TIStructured();
                tis.MachineId = g.Key;
                tis.TimeDict = new List<(int, int)>();
                foreach (var e in g)
                {
                    if(e.mat_id != null && e.op_time != null) tis.TimeDict.Add( ((int)e.op_time, (int)e.mat_id) );
                }
                tis.TimeDict.Sort( (x, y) => x.Item1.CompareTo(y.Item1) );
                StructuredTimes.Add(tis);
            }

        }

        private void LoadCompetitors()
        {
            this.Competitors = new Dictionary<int, Dictionary<int, int>>();

            var groups = from mat in this.Materials
                         join ti in this.Times on mat.Id equals ti.MaterialId into j
                         from submat in j
                         group submat by submat.MaterialId;


            foreach (var g in groups)
            {
                var dict = new Dictionary<int, int>();

                foreach (var e in g)
                {
                    dict.Add(e.MachineId, e.OperationTime);
                }
                this.Competitors.Add(g.Key, dict);
            }
        }


        public List<string> CheckOut()
        {
            List<string> errors = new List<string>();

            var machines_from_times = this.Times.Select(x=>x.MachineId).Distinct();
            var machines_origin = this.Machines.Select(x => x.Id);
            if(machines_from_times.Except(machines_origin).Any()) errors.Add("В times.xlsx есть ID оборудования, которого нет в machine_tools.xlsx");

            var mats_from_times = this.Times.Select(x => x.MaterialId).Distinct();
            var mats_origin = this.Materials.Select(x => x.Id);
            if (mats_from_times.Except(mats_origin).Any()) errors.Add("В times.xlsx есть ID материалов, которых нет в nomenclatures.xlsx");

            var mats_from_parties = this.Parties.Select(x => x.MaterialId).Distinct() ;
            if (mats_from_parties.Except(mats_from_times).Any()) errors.Add("В parties.xlsx есть ID материалов, для которых нет оборудования в times.xlsx");

            if (mats_from_parties.Except(mats_origin).Any()) errors.Add("В parties.xlsx есть ID материалов, которых нет в nomenclatures.xlsx");

            //if(machines_origin.GroupBy(v=>v).Any(g => g.Count() > 1) ) errors.Add("В machine_tools.xlsx имеются повторяющиеся значения идентификаторов");
            //if (mats_origin.GroupBy(v => v).Any(g => g.Count() > 1)) errors.Add("В nomenclatures.xlsx имеются повторяющиеся значения идентификаторов");

            //var parties_id = this.Parties.Select(x => x.Id);
            //if (parties_id.GroupBy(v => v).Any(g => g.Count() > 1)) errors.Add("В parties.xlsx имеются повторяющиеся значения идентификаторов");

            //var pairs_from_times = this.Times.Select(x => (x.MachineId, x.MaterialId) );
            //if (pairs_from_times.GroupBy(v => v).Any(g => g.Count() > 1)) errors.Add("В times.xlsx имеются повторяющиеся пары ID-материала и ID-оборудования");

            return errors;
        }
        
    }
}
