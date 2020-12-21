﻿using System;
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
        public static string path = "";

        public static string direction = "";
        public static string profile = "";
        public static string standard = "";
        public static string protocol = "";
        public static string chair = "";
        public static string subjectName = "";
        public static int creditUnits = 0;
        public static string studyHours = "";
        public static string test = "";
        public static string subjectCompetencies = "";
        public static string subjectIndex = "";
        public static string directionAbbreviation = "";
        public static string startYear = "";
        public static string edForm = "";
        public static Dictionary<string, string> competenciesDic = new Dictionary<string, string>();

        public static int sumLectures = 0;
        public static int sumWorkshops = 0;
        public static string sumIndependentWork = "";

        public static string courseWork = "";
        public static string consulting = "";
        public static string typesOfLessons = "";
        public static string semesters = "";
        public static string courses = "";
        public static Dictionary<string, string> semesterData = new Dictionary<string, string>();
        public static string[] keysForSemesterData = new string[]
        {
            "",
            "$auditoryLessons$",
            "$lectures$",
            "$laboratoryExercises$",
            "$workshops$",
            "$independentWorkBySemester$",
            "$exam$"
        };

        //public static int maxValueOfProgressBar=0;

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
            semesters = "";
            string GradedTest = worksheetPlan.Cells[6][index].Value;
            string testCopy = worksheetPlan.Cells[5][index].Value;
            if (testCopy != null && GradedTest != null)
            {
                if (testCopy.CompareTo(GradedTest) == -1)
                    test = testCopy + GradedTest;
                else
                    test = GradedTest + testCopy;
            }
            else
                test = GradedTest + testCopy;

            string ExamCopy = worksheetPlan.Cells[4][index].Value;
            if (ExamCopy != null && test != null)
            {
                if (ExamCopy.CompareTo(test) == -1)
                    semesters = ExamCopy + test;
                else
                    semesters = test + ExamCopy;
            }
            else
                semesters = test + ExamCopy;
            if (semesters == "")
                for (int i = 18, number = 1; i < 70; i += 7)
                {
                    if (!string.IsNullOrEmpty(worksheetPlan.Cells[i][index].Value))
                        semesters += number;
                    number++;
                }
        }

        public static void FillDictionary(Excel.Worksheet worksheetPlan, int index)
        {

            foreach (var item in semesters)
            {
                int a = Convert.ToInt32(item - '0') - 1;
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
                for (int i = 1; i < 7; i++)
                    if (!semesterData.ContainsKey(keysForSemesterData[i]))
                        semesterData[keysForSemesterData[i]] = "-";
            }
        }
        public static void CreateCourses()
        {
            int a = Convert.ToInt32(Math.Ceiling(Convert.ToDouble(semesters[semesters.Length - 1].ToString()) / 2));
            int b = Convert.ToInt32(Math.Floor(Convert.ToDouble(semesters[0].ToString()) / 2));
            for (int i = b; i < a; i++)
                courses += i + 1 + "/";
            if (semesters.Length==1)
                courses = a.ToString();
            else
                courses = courses.Remove(courses.Length - 1);
        }

        public static void CreateTeats()
        {
            string s = "";
            for (int i = 0, j = 0; i < semesters.Length; i++)
                if (j < test.Length)
                {
                    if (semesters[i] == test[j])
                    {
                        s += "+/";
                        j++;
                    }
                    else
                        s += "-/";
                }
                else
                    s += "-/";
            test = s.Remove(s.Length - 1);
        }

        public static void CreateSemesters()
        {
            string s = "";
            for (int i = 0; i < semesters.Length; i++)
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
                typesOfLessons = list[0] + ", " + list[1] +" и "+list[2];
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

        public static string SelectAbbreviation()
        {
            //Создаем аббревиатуры направлений.
            string directionName = _Excel.worksheetWorkPlanTitlePage.Cells[2][18].Value;
            string abbreviation = "";
            if (directionName.Contains("  "))
                directionName.Replace("  ", " ");
            string[] splittedDirectionName = _Excel.worksheetWorkPlanTitlePage.Cells[2][18].Value.Split(' ');
            if (splittedDirectionName[2] == "Прикладная")
                abbreviation = "ПМ";
            else if (splittedDirectionName[2] == "Информатика")
                abbreviation = "ИВТ";
            else if (splittedDirectionName[2] == "Педагогическое")
                abbreviation = "ПОМИ";
            else
                abbreviation = "МАТ";
            //for (int i = 2; i < directionName.Length; i++)
            //{
            //    if (directionName[i] != "Профиль")
            //    {
            //        if (directionName[i].Length > 1)
            //            abbreviation += Char.ToUpper(directionName[i][0]);
            //    }
            //    else
            //        break;
            //}
            return abbreviation;
        }

        public static void PrepareData(Excel.Worksheet worksheetPlan, Excel.Worksheet worksheetTitle, int index)
        {
            // берём информацию из листа Титул
            subjectName = worksheetPlan.Cells[3][index].Value.Trim(' ');
            string[] separators = new string[] { "Профиль", "Профили", "Направление" };
            var s0 = worksheetTitle.Cells[2][18].Value; //.Split(disciplineSplitArr);
            directionAbbreviation = SelectAbbreviation();
            direction = s0.Split(separators, StringSplitOptions.RemoveEmptyEntries)[0].Trim(' ', ',', ':');
            profile = s0.Split(separators, StringSplitOptions.RemoveEmptyEntries)[1].Trim(' ');
            var s1 = worksheetTitle.Cells[20][31].Value.Split(new string[] { "от" }, StringSplitOptions.RemoveEmptyEntries);
            standard = s1[1].Trim(' ') + " г. " + s1[0].Trim(' ');
            var s2 = worksheetTitle.Cells[1][13].Value.Split(new string[] { "Протокол", "от" }, StringSplitOptions.RemoveEmptyEntries);
            protocol = s2[1].Trim(' ') + " г., " + s2[0].Trim(' ');
            chair = worksheetTitle.Cells[2][26].Value.Trim(' ');
            startYear = worksheetTitle.Cells[20][29].Value.Trim(' ');
            var s3 = worksheetTitle.Cells[1][31].Value.Split(':');
            edForm = s3[1].Trim(' ') + " " + s3[0];
            // берём информацию из листа План
            if (!string.IsNullOrEmpty(worksheetPlan.Cells[8][index].Value))
                creditUnits = int.Parse(worksheetPlan.Cells[8][index].Value.Trim(' '));
            if (!string.IsNullOrEmpty(worksheetPlan.Cells[7][index].Value))
                courseWork = worksheetPlan.Cells[7][index].Value.Trim(' ');
            else
                courseWork = "-";
            studyHours = worksheetPlan.Cells[11][index].Value.Trim(' ') + " час.";
            if (!string.IsNullOrEmpty(worksheetPlan.Cells[14][index].Value))
                sumIndependentWork = worksheetPlan.Cells[14][index].Value.Trim(' ');
            subjectCompetencies = worksheetPlan.Cells[75][index].Value.Trim(' ');
            subjectIndex = worksheetPlan.Cells[2][index].Value.Trim(' ');
            ClearData();
            CreateSemesters(worksheetPlan, index);
            FillDictionary(worksheetPlan, index);
            CreateConsulting();
            CreateCourses();
            CreateTeats();//
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

        private static List<string> SelectCompetencies(Excel.Worksheet worksheet)
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

        public static string RemoveExtraChars(string s)
        {
            //Удаляем лишние символы из названий предметов.
            string str = null;
            foreach (var item in s)
            {
                if (item == ':')
                    str += ' ';
                else
                    str += item;
            }
            return str;
        }

        private void WriteInFile()
        {
            string subjectInPath = "";
            if (subjectName.Contains(':'))
                subjectInPath = RemoveExtraChars(subjectName);
            else
                subjectInPath = subjectName;
            path = folderBrowserDialogChooseFolder.SelectedPath + "\\" + subjectIndex + "_" + subjectInPath + "_" + directionAbbreviation + "_" + startYear;
            var resultList = SelectCompetencies(_Excel.worksheetWorkPlanComp);
            var resultDoc = new _Word();
            resultDoc.path = path;
            var competencies = "\t" + string.Join(";\n\t", resultList) + ".";
            //competenciesDic = CreateCompetenciesDic(_Excel.worksheetWorkPlanPlan);
            resultDoc.FillPattern(competencies);
            //var resultList = SelectCompetencies(worksheet, plan);
            //DocX resultDoc = DocX.Create(path);
            //var resultDoc = DocX.Create(path);
            //var competencies = "\t" + string.Join(";\n\t", resultList) + ".";
            //_Word.CreateWordTemplate(competencies, resultDoc);
            //resultDoc.Save();
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
                        WriteInFile();
                        progressBar1.Value++;
                    }
                }
                labelLoading.Text = "Загрузка завершена";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void buttonOpenFolder_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult res = folderBrowserDialogChooseFolder.ShowDialog();
                if (res == DialogResult.OK)
                {
                    labelNameOfFolder.Text = "Загрузка...";
                    path = folderBrowserDialogChooseFolder.SelectedPath;
                    labelNameOfFolder.Text = path;
                    buttonGenerate.Enabled = true;
                }
                else
                    throw new Exception("Путь не выбран");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        static public int MaxValueOfProgressBar(Excel.Worksheet worksheet)
        {
            int lastRow = TotalSize(worksheet);
            int maxValueOfProgressBar = 0;
            for (int i = 6; i < lastRow; i++)
            {
                if (_Excel.worksheetWorkPlanPlan.Cells[74][i].Value != null || _Excel.worksheetWorkPlanPlan.Cells[10][i].Value != null)
                    maxValueOfProgressBar++;
            }
            return maxValueOfProgressBar;
        }
    }
}
