using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace MetallFactory.Models
{
    public class XlsxRepository: IRepository
    {
        public List<Material> Materials { get; set; }
        public List<Machine> Machines { get; set; }
        public List<Party> Parties { get; set; }
        public List<TimeInfo> Times { get; set; }
        public List<TIStructured> StructuredTimes { get; set; }
        public void Load()
        {
            this.LoadMaterials();
            this.LoadMachines();
            this.LoadParties();
            this.TIRestructure();
        }

        private void LoadMaterials()
        {
            Application xlApp;
            Workbook xlWorkBook;
            Worksheet xlWorkSheet;
            Microsoft.Office.Interop.Excel.Range range;

            Materials = new List<Material>();

            int rCnt;
            int rw = 0;

            xlApp = new Application();

            xlWorkBook = xlApp.Workbooks.Open(@"c:\xlsx_data\nomenclatures.xlsx");

            xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;

            rw = range.Rows.Count;

            for (rCnt = 2; rCnt <= rw; rCnt++)
            {
                Material mat = new Material();
                //get ABC or XYZ
                mat.Id = int.Parse((range.Cells[rCnt, 1] as Microsoft.Office.Interop.Excel.Range).Value2.ToString());
                mat.Name = (string)(range.Cells[rCnt, 2] as Microsoft.Office.Interop.Excel.Range).Value2.ToString();
                Materials.Add(mat);
            }

            //release the resources
            xlWorkBook.Close(true, null, null);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }

        private void LoadMachines()
        {
            Application xlApp;
            Workbook xlWorkBook;
            Worksheet xlWorkSheet;
            Microsoft.Office.Interop.Excel.Range range;

            Machines = new List<Machine>();

            int rCnt;
            int rw = 0;

            xlApp = new Application();

            xlWorkBook = xlApp.Workbooks.Open(@"c:\xlsx_data\machine_tools.xlsx");

            xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;

            rw = range.Rows.Count;

            for (rCnt = 2; rCnt <= rw; rCnt++)
            {
                Machine machine = new Machine();
                //get ABC or XYZ
                machine.Id = int.Parse((range.Cells[rCnt, 1] as Microsoft.Office.Interop.Excel.Range).Value2.ToString());
                machine.Name = (string)(range.Cells[rCnt, 2] as Microsoft.Office.Interop.Excel.Range).Value2.ToString();
                Machines.Add(machine);
            }

            //release the resources
            xlWorkBook.Close(true, null, null);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }

        private void LoadParties()
        {
            Application xlApp;
            Workbook xlWorkBook;
            Worksheet xlWorkSheet;
            Microsoft.Office.Interop.Excel.Range range;

            Parties = new List<Party>();

            int rCnt;
            int rw = 0;

            xlApp = new Application();

            xlWorkBook = xlApp.Workbooks.Open(@"c:\xlsx_data\parties.xlsx");

            xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;

            rw = range.Rows.Count;

            for (rCnt = 2; rCnt <= rw; rCnt++)
            {
                Party party = new Party();
                //get ABC or XYZ
                party.Id = int.Parse((range.Cells[rCnt, 1] as Microsoft.Office.Interop.Excel.Range).Value2.ToString());
                party.MaterialId = int.Parse((range.Cells[rCnt, 2] as Microsoft.Office.Interop.Excel.Range).Value2.ToString());
                Parties.Add(party);
            }

            //release the resources
            xlWorkBook.Close(true, null, null);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }

        private void LoadTimes()
        {
            Application xlApp;
            Workbook xlWorkBook;
            Worksheet xlWorkSheet;
            Microsoft.Office.Interop.Excel.Range range;

            Times = new List<TimeInfo>();

            int rCnt;
            int rw = 0;

            xlApp = new Application();

            xlWorkBook = xlApp.Workbooks.Open(@"c:\xlsx_data\times.xlsx");

            xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;

            rw = range.Rows.Count;

            for (rCnt = 2; rCnt <= rw; rCnt++)
            {
                TimeInfo tinfo = new TimeInfo();
                //get ABC or XYZ
                tinfo.MachineId = int.Parse((range.Cells[rCnt, 1] as Microsoft.Office.Interop.Excel.Range).Value2.ToString());
                tinfo.MaterialId = int.Parse((range.Cells[rCnt, 2] as Microsoft.Office.Interop.Excel.Range).Value2.ToString());
                tinfo.OperationTime = int.Parse((range.Cells[rCnt, 3] as Microsoft.Office.Interop.Excel.Range).Value2.ToString());
                Times.Add(tinfo);
            }

            //release the resources
            xlWorkBook.Close(true, null, null);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }

        private void TIRestructure()
        {
            this.LoadTimes();
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
    }
}
