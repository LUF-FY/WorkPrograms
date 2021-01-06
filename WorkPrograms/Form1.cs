using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Xceed.Words.NET;
using System.IO;
using System.Drawing;

namespace WorkPrograms
{
    public partial class WorkPrograms : Form
    {
        static string filePath = "";

        static string direction = "";
        static string profile = "";
        static string standard = "";
        static string protocol = "";
        static string subjectName = "";
        static int creditUnits = 0;
        static string studyHours = "";
        static string test = "";
        public static string subjectCompetencies = "";
        static string subjectIndex = "";
        static string subjectIndexDecoding = "";
        static string directionAbbreviation = "";
        static string startYear = "";
        static string edForm = "";
        static string interactiveWatch = ""; 
        static int sumLectures = 0;
        static int sumWorkshops = 0;
        public static int sumLaboratoryExercises = 0;
        static string sumIndependentWork = "";
        static string blockName = "";
        static string studyProgram = "";

        static string courseWork = "";
        static string consulting = "";
        static string typesOfLessons = "";
        static List<int> semestersList = new List<int>();
        static string semesters = "";
        static string courses = "";
        static Dictionary<string, string> semesterData = new Dictionary<string, string>();
        static string[] keysForSemesterData = new string[]
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
            semestersList.Clear();
            courses = "";
            test = "";
            sumLectures = 0;
            sumWorkshops = 0;
            sumLaboratoryExercises = 0;
            typesOfLessons = "";
            consulting = "";
            interactiveWatch = "";
        }

        public static void CreateSemesters(Excel.Worksheet worksheetPlan, int index)
        {
            int lastColoumn = TotalSizeColumn(worksheetPlan);
            for (int i = 18, number = 1; i < lastColoumn - 3; i += 7)
            {
                if (!string.IsNullOrEmpty(worksheetPlan.Cells[i][index].Value))
                    semestersList.Add(number);
                number++;
            }
        }

        public static void FillDictionary(Excel.Worksheet worksheetPlan, int index)
        {
            foreach (var item in semestersList)
            {
                int a = item - 1;
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

        public static void CreateIndependentWorkBySemester(Excel.Worksheet worksheetPlan, int index)
        {
            string s = ""; 
            int count = 1;
            int lastColoumn = TotalSizeColumn(worksheetPlan);
            for (int i = 17; i < lastColoumn - 3; i+=7)
            {
                int lec = Convert.ToInt32(worksheetPlan.Cells[i + 2][index].Value);
                int lab = Convert.ToInt32(worksheetPlan.Cells[i + 3][index].Value);
                int pra = Convert.ToInt32(worksheetPlan.Cells[i + 4][index].Value);
                if (semestersList.Contains(count))
                {
                    if (lec + pra + lab != 0)
                        s += (lec + pra + lab) + "/";
                    else
                        s += "-/";        
                }
                count++;
            }
            if (s.Length != 0)
                s = s.Remove(s.Length - 1);
            semesterData[keysForSemesterData[1]] = s;
        }


        public static void CreateCourses(Excel.Worksheet worksheetPlan, int index)
        {
            /*int a = Convert.ToInt32(Math.Ceiling(Convert.ToDouble(semesters[semesters.Length - 1].ToString()) / 2));
            int b = Convert.ToInt32(Math.Floor(Convert.ToDouble(semesters[0].ToString()) / 2));
            for (int i = b; i <= a; i++)
                courses += i + 1 + "/";
            if (semesters.Length==1)
                courses = a.ToString();
            else
                courses = courses.Remove(courses.Length - 1);*/
            int lastColoumn = TotalSizeColumn(worksheetPlan);
            for (int i = 18, number = 1; i < lastColoumn - 3; i += 14)
            {
                if (!string.IsNullOrEmpty(worksheetPlan.Cells[i][index].Value)||
                    !string.IsNullOrEmpty(worksheetPlan.Cells[i+7][index].Value))
                    courses += number+"/";
                number++;
            }
            courses = courses.Remove(courses.Length - 1);
        }

        public static void CreateTeats(Excel.Worksheet worksheetPlan, int index)
        {
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
            string s = "";
            for (int i = 0, j = 0; i < semestersList.Count; i++)
                if (j < test.Length)
                {
                    if (semestersList[i] == test[j])
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
            for (int i = 0; i < semestersList.Count; i++)
                s += semestersList[i] + "/";
            semesters = s.Remove(s.Length - 1);
        }

        public static void CountSumLecturesAndPractices(Excel.Worksheet worksheetPlan, int index)
        {
            int lastColoumn = TotalSizeColumn(worksheetPlan);
            for (int i = 17; i < lastColoumn - 3; i+=7)
            {
                sumLectures += Convert.ToInt32(worksheetPlan.Cells[i + 2][index].Value);
                sumLaboratoryExercises += Convert.ToInt32(worksheetPlan.Cells[i + 3][index].Value);
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
            if (studyProgram == "магистратуры")
                abbreviation = "МАГИ_";
            else if (studyProgram == "аспирантуры")
                abbreviation = "АСПИР_";
            if (directionName.Contains("  "))
                directionName = directionName.Replace("  ", " ");
            string[] splittedDirectionName = _Excel.worksheetWorkPlanTitlePage.Cells[2][18].Value.Split(' ');
            if (splittedDirectionName.Contains("Прикладная"))
                abbreviation += "ПМ";
            else if (profile.Contains("логика"))
                abbreviation += "МЛ";
            else if (profile.Contains("уравнения"))
                abbreviation += "ДУ";
            else if (splittedDirectionName.Contains("Педагогическое"))
                abbreviation += "ПОМИ";
            else if (splittedDirectionName.Contains("Информатика"))
                abbreviation += "ИВТ";
            else
                abbreviation += "МАТ";
            return abbreviation;
        }

        private static string DecodeSubjectIndex(Excel.Worksheet worksheet, int index)
        {
            string subsectionName = "";
            string blockCode1 = "";
            string blockCode2 = "";
            string[] s = subjectIndex.Split('.');
            subjectIndexDecoding = "";
            int i = index;
            if (s[0].ToLower() != blockCode1 || s[1].ToLower() != blockCode2)
            {
                while (!string.IsNullOrEmpty(worksheet.Cells[2][i].Value))
                    i--;
                blockCode1 = s[0].ToLower();
                blockCode2 = s[1].ToLower();
                if (!string.IsNullOrEmpty(worksheet.Cells[1][i - 1].Value))
                {
                    string[] str = worksheet.Cells[1][i - 1].Value.Trim(' ').Split('.');
                    blockName = str[0] + ". " + str[1] + ". ";
                    subsectionName = worksheet.Cells[1][i].Value.Trim(' ');
                }
                else
                    subsectionName = worksheet.Cells[1][i].Value.Trim(' ');
            }
            if (!string.IsNullOrEmpty(blockName) && !string.IsNullOrEmpty(subsectionName))
                subjectIndexDecoding += blockName + subsectionName + ". ";
            if (s.Length > 2)
            {
                if (s[2].ToLower() == "дв")
                    subjectIndexDecoding += "Дисциплины по выбору.";
            }
            return subjectIndexDecoding;
        }

        public static void PrepareData(Excel.Worksheet worksheetPlan, Excel.Worksheet worksheetTitle, int index)
        {
            // берём информацию из листа Титул
            int lastColumn = TotalSizeColumn(worksheetPlan);
            ClearData();
            studyProgram = worksheetTitle.Cells[6][14].Value.Trim(' ').Replace("  ", " ").Split()[2];
            subjectName = worksheetPlan.Cells[3][index].Value.Trim(' ');
            string[] separators = new string[] {"Направленность программы", "Направление подготовки", "Профиль", 
                "Профиль:", "Профили", "Направление"};
            var s0 = worksheetTitle.Cells[2][18].Value; //.Split(disciplineSplitArr);
            direction = s0.Split(separators, StringSplitOptions.RemoveEmptyEntries)[0].Trim(' ', ',', ':');
            try
            {
                profile = "Профиль: " + s0.Split(separators, StringSplitOptions.RemoveEmptyEntries)[1].Trim(' ');
            }
            catch
            {
                profile = "";
            }
            directionAbbreviation = SelectAbbreviation();
            var s1 = worksheetTitle.Cells[20][31].Value.Split(new string[] { "от" }, StringSplitOptions.RemoveEmptyEntries);
            standard = s1[1].Trim(' ') + " г. " + s1[0].Trim(' ');
            var s2 = worksheetTitle.Cells[1][13].Value.Split(new string[] { "Протокол", "от" }, StringSplitOptions.RemoveEmptyEntries);
            protocol = s2[1].Trim(' ') + " г. " + s2[0].Trim(' ');
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
            if (!string.IsNullOrEmpty(worksheetPlan.Cells[16][index].Value))
                interactiveWatch = worksheetPlan.Cells[16][index].Value.Trim(' ');
            subjectCompetencies = worksheetPlan.Cells[lastColumn + 2][index].Value.Trim(' ');
            subjectIndex = worksheetPlan.Cells[2][index].Value.Trim(' ');
            subjectIndexDecoding = DecodeSubjectIndex(worksheetPlan, index);

            CreateSemesters(worksheetPlan, index);
            FillDictionary(worksheetPlan, index);
            CreateIndependentWorkBySemester(worksheetPlan, index);
            CreateConsulting();
            CreateCourses(worksheetPlan, index);
            CreateTeats(worksheetPlan, index);
            CreateSemesters();
            CountSumLecturesAndPractices(worksheetPlan, index);
            CreateTypesOfLessons();
            if(studyProgram== "аспирантуры")
            {
                courses = semesters;
                semesters = "-";
            }
        }
        private static Dictionary<string, string> CreateCompetenciesDic(Excel.Worksheet worksheet)
        {
            // Закидываем в словарь компетенции из листа "Компетенции".
            var dic = new Dictionary<string, string>();
            int lastRow = TotalSizeRow(worksheet);
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
                        resultList.Add($"{item}" + " -" + dic[item]);
                }
            }
            return resultList;
        }

        public static int TotalSizeRow(Excel.Worksheet worksheet)
        {
            // Находим кол-во строк.
            var lastCell = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            return lastCell.Row;
        }

        public static int TotalSizeColumn(Excel.Worksheet worksheet)
        {
            // Находим кол-во столбцов.
            var lastCell = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            return lastCell.Column;
        }


        private void buttonOpenExcel_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult res = openFileDialogSelectFile.ShowDialog();
                if (res == DialogResult.OK)
                {
                    SelectFile.SelectExcelWorkPlanFile(openFileDialogSelectFile, labelNameOfWorkPlanFile);
                    buttonOpenFolder.Enabled = true;
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
                if (item == ':' || item == '\\' || item == '|' || item == '/' || 
                        item == '*' || item == '?' || item == '"' || item == '>' || item == '<')
                    str += ' ';
                else
                    str += item;
            }
            return str;
        }

        private void WriteInFile()
        {
            string subjectInPath = RemoveExtraChars(subjectName);
            filePath = folderBrowserDialogChooseFolder.SelectedPath + "\\" + subjectIndex + "_" + subjectInPath + "_" + directionAbbreviation + "_" + startYear;
            var resultList = SelectCompetencies(_Excel.worksheetWorkPlanComp);
            var resultDoc = new _Word();
            resultDoc.path = filePath;
            var competencies = "\t" + string.Join(";\n\t", resultList) + ".";
            var competenciesDic = CreateCompetenciesDic(_Excel.worksheetWorkPlanComp); 
            string[] replaceableStrings = new string[]
            {
                subjectName, direction, profile,
                standard, protocol, creditUnits.ToString(),
                studyHours, courses, semesters, sumIndependentWork.ToString(),
                typesOfLessons, test, consulting, courseWork,
                competencies, edForm, sumLectures.ToString(), sumWorkshops.ToString(), interactiveWatch,
                subjectIndex, subjectIndexDecoding
            };
            string[] namesOfReplaceableStrings = new string[]
            {
                nameof(subjectName), nameof(direction), nameof(profile),
                nameof(standard), nameof(protocol),nameof(creditUnits), nameof(studyHours),
                nameof(courses), nameof(semesters), nameof(sumIndependentWork),nameof(typesOfLessons),
                nameof(test), nameof(consulting), nameof(courseWork), nameof(competencies), 
                nameof(edForm), nameof(sumLectures), nameof(sumWorkshops), nameof(interactiveWatch),
                nameof(subjectIndex), nameof(subjectIndexDecoding)
            };
            bool isInteractiveWatch = true;
            if (string.IsNullOrEmpty(interactiveWatch))
                isInteractiveWatch = false;
            resultDoc.FillPattern(competenciesDic, replaceableStrings, namesOfReplaceableStrings, semesterData, isInteractiveWatch);
        }

        private void buttonGenerate_Click(object sender, EventArgs e)
        {
            //Создаем файлы .            
            try
            {
                labelLoading.Visible = true;
                labelLoading.Text = "Загрузка...";             
                int lastRow = TotalSizeRow(_Excel.worksheetWorkPlanPlan);
                int lastColumn = TotalSizeColumn(_Excel.worksheetWorkPlanPlan);
                progressBar1.Maximum = MaxValueOfProgressBar(_Excel.worksheetWorkPlanPlan);
                for (int i = 6; i <= lastRow; i++)
                {
                    if (_Excel.worksheetWorkPlanPlan.Cells[lastColumn+1][i].Value != null || _Excel.worksheetWorkPlanPlan.Cells[10][i].Value != null)
                    {
                        PrepareData(_Excel.worksheetWorkPlanPlan, _Excel.worksheetWorkPlanTitlePage, i);
                        WriteInFile();
                        progressBar1.Value++;
                    }
                }
                labelLoading.Text = "Загрузка завершена";
                MessageBox.Show("Загрузка завершена");
                Reset();
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
                    filePath = folderBrowserDialogChooseFolder.SelectedPath;
                    labelNameOfFolder.Text = filePath;
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
            int lastRow = TotalSizeRow(worksheet);
            int lastColumn = TotalSizeColumn(worksheet);
            int maxValueOfProgressBar = 0;
            for (int i = 6; i <= lastRow; i++)
            {
                if (_Excel.worksheetWorkPlanPlan.Cells[lastColumn+1][i].Value != null || _Excel.worksheetWorkPlanPlan.Cells[10][i].Value != null)
                    maxValueOfProgressBar++;
            }
            return maxValueOfProgressBar;
        }

        void Reset()
        {
            progressBar1.Value = 0;
            progressBar1.Maximum = 0;
            labelLoading.Text = "Ожидание";
            labelNameOfFolder.Text = "Папка не выбрана";
            labelNameOfWorkPlanFile.Text = "Файл не выбран";
        }
    }
}
