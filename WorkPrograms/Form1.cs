using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Xceed.Words.NET;
using System.IO;

namespace WorkPrograms
{
    public partial class WorkPrograms : Form
    {
        
        public static string direction = "";
        public static string profile = "";
        public static string standard = "";
        public static string protocol = "";
        public static string chair = "";
        public static string subjectName = "";
        public static int creditUnits = 0;
        public static int studyHours = 0;
        public static int independentWork = 0;
        public static string test = "";

        public WorkPrograms()
        {
            InitializeComponent();
        }
        public static void PrepareData(Excel.Worksheet worksheetPlan, Excel.Worksheet worksheetTitle, int index)
        {
            // берём информацию из листа Титул
            subjectName = worksheetPlan.Cells[3][index].Value.Trim(' ');
            var s0 = worksheetTitle.Cells[2][18].Value.Split(new string[] { "Профиль", "Профили", "Направление" });
            direction = s0[0].Trim(' ');
            profile = s0[1].Trim(' ');
            var s1 = worksheetTitle.Cells[20][31].Value.Split("от");
            standard = s1[1].Trim(' ') + " г. " + s1[0].Trim(' ');
            var s2 = worksheetTitle.Cells[1][13].Value.Split("от");
            protocol = s2[1].Trim(' ') + " г., " + s2[0].Trim(' ');
            chair = worksheetTitle.Cells[2][26].Value.Trim(' ');
            // берём информацию из листа План
            if (!string.IsNullOrEmpty(worksheetPlan.Cells[8][index].Value))
                creditUnits = int.Parse(worksheetPlan.Cells[8][index].Value);
            studyHours = int.Parse(worksheetPlan.Cells[11][index].Value);
            independentWork = int.Parse(worksheetPlan.Cells[14][index].Value);
            test = worksheetPlan.Cells[5][index].Value;

        }
        private static Dictionary<string, string> CreateCompetenciesDic(Excel.Worksheet worksheet)
        {
            // Закидываем в словарь компетенции из листа "Компетенции".
            var dic = new Dictionary<string, string>();
            int lastRow = TotalSize(worksheet);
            for (int i = 3; i < lastRow; i++)
            {
                if (!string.IsNullOrEmpty(worksheet.Cells[2][i].Value))
                {
                    string key = worksheet.Cells[2][i].Value;
                    dic[key] = worksheet.Cells[4][i].Value;
                }
            }
            return dic;
        }

        private List<string> SelectCompetencies(Excel.Worksheet worksheet, Excel.Worksheet worksheet2)
        {
            // Ищем в листе "Компетенции" нужные компетенции и закидываем в список.
            var resultList = new List<string>();
            var dic = CreateCompetenciesDic(worksheet);

            var competenciesList = subjectCompetencies.Split(';', ' ').ToList();
            foreach (var item in competenciesList)
            {
                if (!string.IsNullOrEmpty(item))
                {
                    if (dic.ContainsKey(item))
                        resultList.Add("-" + dic[item] + " " + $"({item})");
                }
            }
            return resultList;
        }

        /*что делает Стас
         *  var competencies = "\t" + string.Join(";\n\t", resultList) + ".";
         * 
         * 
        */
        public static int TotalSize(Excel.Worksheet worksheet)
        {
            // Находим кол-во строк.
            var lastCell = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            return lastCell.Row;
        }

        private void buttonOpenExcel_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult res = openFileDialogSelectFile.ShowDialog();
                if (res == DialogResult.OK)
                {
                    SelectFile.SelectExcelWorkPlanFile(openFileDialogSelectFile, labelNameOfWorkPlanFile);                    
                }
                else
                    throw new Exception("Файл не выбран");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void buttonGenerate_Click(object sender, EventArgs e)
        {
            //Создаем файлы .            
            try
            {
                
                int lastRow = TotalSize(_Excel.worksheetWorkPlanPlan);
                labelLoading.Text = "Загрузка...";                
                for (int i = 6; i <= lastRow; i++)
                {
                    if (_Excel.worksheetWorkPlanPlan.Cells[74][i].Value != null || _Excel.worksheetWorkPlanPlan.Cells[10][i].Value != null)
                    {
                        PrepareData(_Excel.worksheetWorkPlanPlan, i);
                        //WriteCompetencyInFile(_Excel.worksheetWorkPlanComp, _Excel.worksheetWorkPlanPlan);
                        //isExam = false;
                        //isTest = false;
                    }
                }
                labelLoading.Text = "Загрузка завершена";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
