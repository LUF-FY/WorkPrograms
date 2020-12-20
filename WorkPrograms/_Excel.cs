using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace WorkPrograms
{
    class _Excel
    {
        public static Excel.Application xlApp = null;
        public static Excel.Workbook xlWorkPlan = null;
        public static Excel.Worksheet worksheetWorkPlanComp = null;
        public static Excel.Worksheet worksheetWorkPlanPlan = null;
        public static Excel.Worksheet worksheetWorkPlanTitlePage = null;
        public static void ClearExcel()
        {
            Marshal.ReleaseComObject(xlWorkPlan);
            Marshal.ReleaseComObject(worksheetWorkPlanComp);
            Marshal.ReleaseComObject(worksheetWorkPlanPlan);
            Marshal.ReleaseComObject(xlApp);
        }
    }
}
