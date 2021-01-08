using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Xceed.Words.NET;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace WorkPrograms
{
    /* Заполняем ComboBox предметов и выбираем файл с расширением xls.*/

    class SelectFile
    {
        public async static void SelectExcelWorkPlanFile(OpenFileDialog SelectFile, Label NameOfExcelFile)/*, ComboBox comboBox1*/
        {
            await Task.Run(() => { 
            NameOfExcelFile.Text = "Загрузка...";
            string xlPath = SelectFile.FileName;
            _Excel.xlApp = new Excel.Application();
            _Excel.xlWorkPlan = _Excel.xlApp.Workbooks.Open(xlPath);
            _Excel.worksheetWorkPlanComp = _Excel.xlWorkPlan.Worksheets["Компетенции"];
            _Excel.worksheetWorkPlanPlan = _Excel.xlWorkPlan.Worksheets["План"];
            _Excel.worksheetWorkPlanTitlePage = _Excel.xlWorkPlan.Worksheets["Титул"];
            NameOfExcelFile.Text = Path.GetFileNameWithoutExtension(xlPath);
        });
        }
    }
}