using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
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

        public static void SelectExcelWorkPlanFile(OpenFileDialog SelectFile, Label NameOfExcelFile)/*, ComboBox comboBox1*/
        {
            NameOfExcelFile.Text = "Загрузка...";
            string xlPath = SelectFile.FileName;
            xlApp = new Excel.Application();
            xlWorkPlan = xlApp.Workbooks.Open(xlPath);
            worksheetWorkPlanComp = xlWorkPlan.Worksheets["Компетенции"];
            worksheetWorkPlanPlan = xlWorkPlan.Worksheets["План"];
            worksheetWorkPlanTitlePage = xlWorkPlan.Worksheets["Титул"];
            NameOfExcelFile.Text = Path.GetFileNameWithoutExtension(xlPath);
        }

        public static void ClearExcel()
        {
            Marshal.ReleaseComObject(xlWorkPlan);
            Marshal.ReleaseComObject(worksheetWorkPlanComp);
            Marshal.ReleaseComObject(worksheetWorkPlanPlan);
            Marshal.ReleaseComObject(xlApp);
        }
    }
}
