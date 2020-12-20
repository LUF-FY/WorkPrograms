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
        public static string test = "";
        public static string subjectCompetencies = "";

        public static int sumLectures = 0;
        public static int sumWorkshops = 0;
        public static int sumIndependentWork = 0;

        public static string courseWork = "";
        public static string consulting = "";
        public static string typesOfLessons = "";
        public static string semesters = "";
        public static string courses = "";
        public static Dictionary<string, string> semesterData = new Dictionary<string, string>();
        public static string[] keysForSemesterData = new string[]
        {
            "",
            "auditoryLessons",
            "lectures",
            "laboratoryExercises",
            "workshops",
            "independentWorkBySemester",
            "exam"
        };

        public WorkPrograms()
        {
            InitializeComponent();
        }

        public static void ClearData()
        {
            semesterData.Clear();
            semesters = "";
            courses = "";
            test = "";
            sumLectures = 0;
            sumWorkshops = 0;
            typesOfLessons = "";
            consulting = "";
        }

        public static void CreateSemesters(Excel.Worksheet worksheetPlan, int index)
        {
            string GradedTest = worksheetPlan.Cells[6][index].Value;
            string testCopy = worksheetPlan.Cells[5][index].Value;
            if (testCopy != null && GradedTest != null)
            {
                if (testCopy.CompareTo(GradedTest) == 1)
                    test = testCopy + GradedTest;
                else
                    test = GradedTest + testCopy;
            }

            string ExamCopy = worksheetPlan.Cells[4][index].Value;
            if (ExamCopy != null && test != null)
            {
                if (ExamCopy.CompareTo(test) == 1)
                    semesters = ExamCopy + test;
                else
                    semesters = test + ExamCopy;
            }
        }

        public static void FillDictionary(Excel.Worksheet worksheetPlan, int index)
        {

            foreach (var item in semesters)
            {
                int a = Convert.ToInt32(item);
                for (int i = 1; i < 7; i++)
                {
                    string s3 = worksheetPlan.Cells[(a * 7 + 17 + i)][index].Value;
                    if (s3 != null)
                    {
                        if (!semesterData.ContainsKey(keysForSemesterData[i]))
                            semesterData.Add(keysForSemesterData[i], s3);
                        else
                            semesterData[keysForSemesterData[i]] += "/" + s3;
                    }
                    else if (i == 6)
                    {
                        if (!semesterData.ContainsKey(keysForSemesterData[i]))
                            semesterData.Add(keysForSemesterData[i], "-");
                        else
                            semesterData[keysForSemesterData[i]] += "/-";
                    }
                }
            }
        }
        public static void CreateCourses()
        {
            int a = Convert.ToInt32(Math.Ceiling(Convert.ToDouble(semesters[semesters.Length - 1]) / 2));
            for (int i = 0; i < a; i++)
                courses += i + 1 + "/";
            courses.Trim('/');
        }

        public static void CreateTeats()
        {
            string s = "";
            for (int i = 0, j = 0; i < semesters.Length - 1; i++)
                if (semesters[i] == test[j])
                {
                    s += "+/";
                    j++;
                }
                else
                    s += "-/";
            test = s.Remove(s.Length - 1);
        }

        public static void CreateSemesters()
        {
            string s = "";
            for (int i = 0; i < semesters.Length - 1; i++)
                s += semesters[i] + "/";
            semesters = s.Remove(s.Length - 1);
        }

        public static void CountSumLecturesAndPractices(Excel.Worksheet worksheetPlan, int index)
        {
            for (int i = 17; i < 73; i+=7)
            {
                sumLectures += Convert.ToInt32(worksheetPlan.Cells[i+2][index].Value);
                sumWorkshops += Convert.ToInt32(worksheetPlan.Cells[i + 4][index].Value);
            }        
        }

        public static void CreateTypesOfLessons() 
        {
            var list = new List<string>();
            if (sumLectures!=0)
                list.Add("лекционных");
            if (sumWorkshops!=0)
                list.Add("практических");
            if (semesterData.ContainsKey(keysForSemesterData[3]))
                list.Add("лабораторных");
            if (list.Count==1)
                typesOfLessons = list[0];
            else if (list.Count == 2)
                typesOfLessons = list[0] + ", " + list[1];
            else if (list.Count == 3)
                typesOfLessons = list[0] + ", " + list[1] +"и "+list[2];
        }

        public static void CreateConsulting()
        {
            string[] s = semesterData[keysForSemesterData[6]].Split('/');
            for (int i = 0; i < s.Length; i++)
            {
                if (s[i]!="-")
                    consulting += "+/";
                else
                    consulting += "-/";
            }
            consulting = consulting.Remove(consulting.Length - 1);
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
            if (!string.IsNullOrEmpty(worksheetPlan.Cells[7][index].Value))
                courseWork = worksheetPlan.Cells[7][index].Value;
            studyHours = int.Parse(worksheetPlan.Cells[11][index].Value);
            sumIndependentWork = int.Parse(worksheetPlan.Cells[14][index].Value);
            subjectCompetencies = worksheetPlan.Cells[75][index].Value.Trim(' ');
            ClearData();
            CreateSemesters(worksheetPlan, index);
            FillDictionary(worksheetPlan, index);
            CreateConsulting();
            CreateCourses();
            CreateTeats();
            CreateSemesters();
            CountSumLecturesAndPractices(worksheetPlan, index);
            CreateTypesOfLessons();
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

        private string SelectCompetencies(Excel.Worksheet worksheet, Excel.Worksheet worksheet2)
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
            var competencies = "\t" + string.Join(";\n\t", resultList) + ".";
            return competencies;
        }

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
                        PrepareData(_Excel.worksheetWorkPlanPlan, _Excel.worksheetWorkPlanTitlePage, i);
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
