using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Xceed.Words.NET;
using System.IO;
using System.Drawing;
using System.Threading.Tasks;

namespace WorkPrograms
{
    public partial class WorkPrograms : Form
    {
        string filePath = "";
        public string subjectCompetencies = "";
        //public int sumLaboratoryExercises = 0;
        string blockName = "";

        public WorkPrograms()
        {
            InitializeComponent();
        }

        string GetStudyProgram(Excel.Worksheet worksheetTitle)
        {
            var s = worksheetTitle.Cells[6][14].Value;
            return s.Trim(' ').Replace("  ", " ").Split()[2];
        }

        /// <summary>
        /// Получает направление и профиль
        /// </summary>
        /// <param name="worksheetTitle">Титульный лист</param>
        /// <returns>Массив {напрвление, профиль}</returns>
        string[] GetDirectionAndProfile(Excel.Worksheet worksheetTitle)
        {
            string[] separators = new string[] {"Направленность программы", "Направление подготовки", "Профиль",
                "Профиль:", "Профили", "Направление"}; //разделители направления и профиля
            var directionAndProfile = worksheetTitle.Cells[2][18].Value.Split(separators, StringSplitOptions.RemoveEmptyEntries); //Сплит по разделителям
            var direction = directionAndProfile[0].Trim(' ', ',', ':'); //Получить направление
            string profile = "";
            if (directionAndProfile.Length > 1)
                profile = "Профиль: " + directionAndProfile[1].Trim(' '); //Получить профиль, если он есть
            return new string[] { direction, profile };
        }

        string GetStandart(Excel.Worksheet worksheetTitle)
        {
            var s = worksheetTitle.Cells[20][31].Value.Split(new string[] { "от" }, StringSplitOptions.RemoveEmptyEntries);
            return s[1].Trim(' ') + " г. " + s[0].Trim(' ');
        }

        string GetProtocol(Excel.Worksheet worksheetTitle)
        {
            var s = worksheetTitle.Cells[1][13].Value.Split(new string[] { "Протокол", "от" }, StringSplitOptions.RemoveEmptyEntries);
            return s[1].Trim(' ') + " г. " + s[0].Trim(' ');
        }

        string GetEdForm(Excel.Worksheet worksheetTitle)
        {
            var s = worksheetTitle.Cells[1][31].Value.Split(':');
            return s[1].Trim(' ') + " " + s[0];
        }

        string GetDirectionAbbreviation(Excel.Worksheet worksheetTitle, Dictionary<string, string> dic)
        {
            //Создаем аббревиатуры направлений.
            string directionName = worksheetTitle.Cells[2][18].Value;
            string abbreviation = "";
            if (dic["$studyProgram$"] == "магистратуры")
                abbreviation = "МАГИ_";
            else if (dic["$studyProgram$"] == "аспирантуры")
                abbreviation = "АСПИР_";
            if (directionName.Contains("  "))
                directionName = directionName.Replace("  ", " ");
            string[] splittedDirectionName = worksheetTitle.Cells[2][18].Value.Split(' ');
            if (splittedDirectionName.Contains("Прикладная"))
                abbreviation += "ПМ";
            else if (dic["$profile$"].Contains("логика"))
                abbreviation += "МЛ";
            else if (dic["$profile$"].Contains("уравнения"))
                abbreviation += "ДУ";
            else if (splittedDirectionName.Contains("Педагогическое"))
                abbreviation += "ПОМИ";
            else if (splittedDirectionName.Contains("Информатика"))
                abbreviation += "ИВТ";
            else
                abbreviation += "МАТ";
            return abbreviation;
        }

        Dictionary<string, string> PrepareDataFromSheetTitle(Excel.Worksheet worksheetTitle)
        {
            var dic = new Dictionary<string, string>();
            dic.Add("$studyProgram$", GetStudyProgram(worksheetTitle));
            dic.Add("$direction$", GetDirectionAndProfile(worksheetTitle)[0]);
            dic.Add("$profile$", GetDirectionAndProfile(worksheetTitle)[1]);
            dic.Add("$standart$", GetStandart(worksheetTitle));
            dic.Add("$protocol$", GetProtocol(worksheetTitle));
            dic.Add("$edForm$", GetEdForm(worksheetTitle));
            dic.Add("$directionAbbreviation$", GetDirectionAbbreviation(worksheetTitle, dic));
            dic.Add("$director$", "А.М. Дигурова");
            dic.Add("$position$", "Проректор по УР");
            if (dic["$studyProgram$"] == "аспирантуры")
            {
                dic["$director$"] = "Б.В. Туаева";
                dic["$position$"] = "Проректор по научной деятельности";
            }
            else if (dic["$studyProgram$"] == "магистратуры")
            {
                dic["$director$"] = "Л.А. Агузарова";
                dic["$position$"] = "Первый проректор";
            }
            return dic;
        }



        string GetSubjectName(Excel.Worksheet worksheetPlan, int index) => worksheetPlan.Cells[3][index].Value.Trim(' ');

        string GetCreditUnits(Excel.Worksheet worksheetPlan, int index)
        {
            if (!string.IsNullOrEmpty(worksheetPlan.Cells[8][index].Value))
                return worksheetPlan.Cells[8][index].Value.Trim(' ');
            else
                return "";
        }

        string GetStudyHours(Excel.Worksheet worksheetPlan, int index) => worksheetPlan.Cells[11][index].Value.Trim(' ') + " час.";

        string GetSumIndependentWork(Excel.Worksheet worksheetPlan, int index)
        {
            if (!string.IsNullOrEmpty(worksheetPlan.Cells[14][index].Value))
                return worksheetPlan.Cells[14][index].Value.Trim(' ');
            else
                return "";
        }

        string GetInteractiveWatch(Excel.Worksheet worksheetPlan, int index)
        {
            if (!string.IsNullOrEmpty(worksheetPlan.Cells[16][index].Value))
                return worksheetPlan.Cells[16][index].Value.Trim(' ');
            else
                return "";
        }

        string GetSubjectCompetencies(Excel.Worksheet worksheetPlan, int index, int lastColumn) => worksheetPlan.Cells[lastColumn + 2][index].Value.Trim(' ');

        string GetSubgectIndex(Excel.Worksheet worksheetPlan, int index) => worksheetPlan.Cells[2][index].Value.Trim(' ');

        string GetCourseWork(Excel.Worksheet worksheetPlan, int index)
        {
            if (!string.IsNullOrEmpty(worksheetPlan.Cells[7][index].Value))
                return worksheetPlan.Cells[7][index].Value.Trim(' ');
            else
                return "-";
        }

        string DecodeSubjectIndex(Excel.Worksheet worksheet, int index, string subjectIndex)
        {
            string subsectionName = "";
            string blockCode1 = "";
            string blockCode2 = "";
            string[] s = subjectIndex.Split('.');
            string subjectIndexDecoding = "";
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

        List<int> CreateSemesters(Excel.Worksheet worksheetPlan, int index, int lastColumn)
        {
            var list = new List<int>();
            for (int i = 18, number = 1; i < lastColumn - 3; i += 7)
            {
                if (!string.IsNullOrEmpty(worksheetPlan.Cells[i][index].Value))
                    list.Add(number);
                number++;
            }
            return list;
        }

        Dictionary<string, string> FillDictionary(Excel.Worksheet worksheetPlan, int index, List<int> semestersList, string[] keys)
        {
            var dic = new Dictionary<string, string>();
            foreach (var item in semestersList)
            {
                int a = item - 1;
                for (int i = 0; i < 6; i++)
                {
                    string s3 = worksheetPlan.Cells[(a * 7 + 17 + i + 1)][index].Value;
                    if (s3 != null)
                    {
                        if (!dic.ContainsKey(keys[i]))
                            dic.Add(keys[i], s3);
                        else
                            dic[keys[i]] += "/" + s3;
                    }
                    else if (i == 5)
                    {
                        if (!dic.ContainsKey(keys[i]))
                            dic.Add(keys[i], "-");
                        else
                            dic[keys[i]] += "/-";
                    }
                }
                for (int i = 0; i < 6; i++)
                    if (!dic.ContainsKey(keys[i]))
                        dic[keys[i]] = "-";
            }
            return dic;
        }

        void GetDataFromSemesters(Dictionary<string, string> dic, Excel.Worksheet worksheetPlan, int index, List<int> semestersList)
        {
            var keysTemporaryDic = new string[] 
            { 
                "$auditoryLessons$", 
                "$lectures$", 
                "$laboratoryExercises$",
                "$workshops$", 
                "$independentWorkBySemester$", 
                "$exam$" 
            };
            var temporaryDic = FillDictionary(worksheetPlan, index, semestersList, keysTemporaryDic);
            foreach (var item in temporaryDic)
                dic.Add(item.Key, item.Value);
        }

        string GetIndependentWorkBySemester(Excel.Worksheet worksheetPlan, int index, int lastColumn, List<int> semestersList)
        {
            string s = "";
            int count = 1;
            for (int i = 17; i < lastColumn - 3; i += 7)
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
            return s;
        }

        string CreateConsulting(string exam)
        {
            var ss = "";
            string[] s = exam.Split('/');
            for (int i = 0; i < s.Length; i++)
            {
                if (s[i] != "-")
                    ss += "+/";
                else
                    ss += "-/";
            }
            if (ss.Length != 0)
                ss = ss.Remove(ss.Length - 1);
            return ss;
        }

        string CreateCourses(Excel.Worksheet worksheetPlan, int index, int lastColumn)
        {
            string s = "";
            for (int i = 18, number = 1; i < lastColumn - 3; i += 14)
            {
                if (!string.IsNullOrEmpty(worksheetPlan.Cells[i][index].Value) ||
                    !string.IsNullOrEmpty(worksheetPlan.Cells[i + 7][index].Value))
                    s += number + "/";
                number++;
            }
            if (s.Length != 0)
                s = s.Remove(s.Length - 1);
            return s;
        }

        string CreateTests(Excel.Worksheet worksheetPlan, int index, List<int> semestersList)
        {
            string ss = "";
            string GradedTest = worksheetPlan.Cells[6][index].Value;
            string testCopy = worksheetPlan.Cells[5][index].Value;
            if (testCopy != null && GradedTest != null)
            {
                if (testCopy.CompareTo(GradedTest) == -1)
                    ss = testCopy + GradedTest;
                else
                    ss = GradedTest + testCopy;
            }
            else
                ss = GradedTest + testCopy;
            string s = "";
            for (int i = 0, j = 0; i < semestersList.Count; i++)
                if (j < ss.Length)
                {
                    if (semestersList[i].ToString() == ss[j].ToString())
                    {
                        s += "+/";
                        j++;
                    }
                    else
                        s += "-/";
                }
                else
                    s += "-/";
            if (s.Length != 0)
                s = s.Remove(s.Length - 1);
            return s;
        }

        string CreateSemesters(List<int> semestersList)
        {
            string s = "";
            for (int i = 0; i < semestersList.Count; i++)
                s += semestersList[i] + "/";
            if (s.Length != 0)
                s = s.Remove(s.Length - 1);
            return s;
        }

        string[] CountSumLecturesAndPractices(Excel.Worksheet worksheetPlan, int index, int lastColumn)
        {
            int sumLectures = 0;
            int sumLaboratoryExercises = 0;
            int sumWorkshops = 0;
            for (int i = 17; i < lastColumn - 3; i += 7)
            {
                sumLectures += Convert.ToInt32(worksheetPlan.Cells[i + 2][index].Value);
                sumLaboratoryExercises += Convert.ToInt32(worksheetPlan.Cells[i + 3][index].Value);
                sumWorkshops += Convert.ToInt32(worksheetPlan.Cells[i + 4][index].Value);
            }
            return new string[] { sumLectures.ToString(), sumLaboratoryExercises.ToString(), sumWorkshops.ToString() };
        }

        string CreateTypesOfLessons(string sumLectures, string sumWorkshops, string sumLaboratoryExercises)
        {
            string s = "";
            var list = new List<string>();
            if (sumLectures != "0")
                list.Add("лекционных");
            if (sumWorkshops != "0")
                list.Add("практических");
            if (sumLaboratoryExercises != "0")
                list.Add("лабораторных");
            if (list.Count == 1)
                s = list[0];
            else if (list.Count == 2)
                s = list[0] + " и " + list[1];
            else if (list.Count == 3)
                s = list[0] + ", " + list[1] + " и " + list[2];
            if (s.Length != 0)
                s = s.Remove(s.Length - 1);
            return s;
        }

        Dictionary<string, string> PrepareDataFromSheetPlan
            (Excel.Worksheet worksheetPlan, int index, int lastColumn, Dictionary<string, string> titleDic)
        {
            //dic.Add("$$",);
            var dic = new Dictionary<string, string>();
            dic.Add("$subjectName$", GetSubjectName(worksheetPlan, index));
            dic.Add("$creditUnits$", GetCreditUnits(worksheetPlan, index));
            dic.Add("$studyHours$", GetStudyHours(worksheetPlan, index));
            dic.Add("$sumIndependentWork$", GetSumIndependentWork(worksheetPlan, index));
            dic.Add("$interactiveWatch$", GetInteractiveWatch(worksheetPlan, index));
            dic.Add("$subjectCompetencies$", GetSubjectCompetencies(worksheetPlan, index, lastColumn));
            dic.Add("$subjectIndex$", GetSubgectIndex(worksheetPlan, index));
            dic.Add("$courseWork$", GetCourseWork(worksheetPlan, index));
            dic.Add("$subjectIndexDecoding$", DecodeSubjectIndex(worksheetPlan, index, dic["$subjectIndex$"]));
            var semestersList = CreateSemesters(worksheetPlan, index, lastColumn);
            GetDataFromSemesters(dic, worksheetPlan, index, semestersList);
            dic["$independentWorkBySemester$"] = GetIndependentWorkBySemester(worksheetPlan, index, lastColumn, semestersList);
            dic.Add("$consulting$", CreateConsulting(dic["exam"]));
            dic.Add("$courses$", CreateCourses(worksheetPlan, index, lastColumn));
            dic.Add("$test$", CreateTests(worksheetPlan, index, semestersList));
            dic.Add("$semesters$", CreateSemesters(semestersList));
            var sumLecturesAndPractices = CountSumLecturesAndPractices(worksheetPlan, index, lastColumn);
            dic.Add("$sumLectures$", sumLecturesAndPractices[0]);
            dic.Add("$sumLabs$", sumLecturesAndPractices[1]);
            dic.Add("$sumWorkshops$", sumLecturesAndPractices[2]);
            dic.Add("$typesOfLessons$", CreateTypesOfLessons(dic["$sumLectures$"], dic["$sumLabs$"], dic["$sumWorkshops$"]));
            if (titleDic["$studyProgram$"] == "аспирантуры")
            {
                dic["$courses$"] = dic["$semesters$"];
                dic["$semesters$"] = "-";
            }
            return dic;
        }

        private static Dictionary<string, string> CreateCompetenciesDic(Excel.Worksheet worksheet)
        {
            // Закидываем в словарь компетенции из листа "Компетенции".
            var dic = new Dictionary<string, string>();
            int lastRow = TotalSize(worksheet)[0];
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

        private List<string> SelectCompetencies(Excel.Worksheet worksheet)
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
        

        /// <summary>
        /// Находит последние строку и столбец в листе Excel файла
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns>{последняя строка, последний столбец }</returns>
        public static int[] TotalSize(Excel.Worksheet worksheet)
        {
            // Находим кол-во строк.
            var lastCell = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            return new int[] { lastCell.Row,  lastCell.Column};
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
            //string startYear = worksheetTitle.Cells[20][29].Value.Trim(' '); // Это в название
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
                competencies, edForm, sumLectures.ToString(), sumWorkshops.ToString(), sumLabs.ToString(),
                interactiveWatch, subjectIndex, subjectIndexDecoding, director, position, studyProgram
            };
            string[] namesOfReplaceableStrings = new string[]
            {
                nameof(subjectName), nameof(direction), nameof(profile),
                nameof(standard), nameof(protocol),nameof(creditUnits), nameof(studyHours),
                nameof(courses), nameof(semesters), nameof(sumIndependentWork),nameof(typesOfLessons),
                nameof(test), nameof(consulting), nameof(courseWork), nameof(competencies),
                nameof(edForm), nameof(sumLectures), nameof(sumWorkshops), nameof(sumLabs), nameof(interactiveWatch),
                nameof(subjectIndex), nameof(subjectIndexDecoding), nameof(director), nameof(position), nameof(studyProgram)
            };
            bool isInteractiveWatch = true;
            if (string.IsNullOrEmpty(interactiveWatch))
                isInteractiveWatch = false;
            resultDoc.FillPattern(competenciesDic, replaceableStrings, namesOfReplaceableStrings, semesterData, isInteractiveWatch);
        }

        void SelectExcelFile(OpenFileDialog SelectFile)
        {
            labelNameOfWorkPlanFile.Text = "Загрузка...";
            string xlPath = SelectFile.FileName;
            _Excel.SelectExcelWorkPlanFile(xlPath);
            labelNameOfWorkPlanFile.Text = Path.GetFileNameWithoutExtension(xlPath);
        }


        /// <summary>
        /// Выбор Excel файла с учебным планом, и выбор нужных страниц 
        /// </summary>
        private void buttonOpenExcel_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult res = openFileDialogSelectFile.ShowDialog(); //Выбор файла 
                if (res == DialogResult.OK) //Если файл выбран
                {
                    SelectExcelFile(openFileDialogSelectFile); //Выбор нужных листов
                    buttonOpenFolder.Enabled = true; //Разблокировка кнопки выбора папки
                }
                else
                    throw new Exception("Файл не выбран");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Выбор папки создания шаблонов рабочих программ, и сохранение путя
        /// </summary>
        private void buttonOpenFolder_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult res = folderBrowserDialogChooseFolder.ShowDialog(); //Выбор папки
                if (res == DialogResult.OK) //Если папка выбрана
                {
                    labelNameOfFolder.Text = "Загрузка..."; // изменение лейбла состояния
                    filePath = folderBrowserDialogChooseFolder.SelectedPath; // сохранение путя
                    labelNameOfFolder.Text = filePath; //изменение лейбла на путь
                    buttonGenerate.Enabled = true; //Разблокировка кнопки для свормировывания шаблонов
                }
                else
                    throw new Exception("Путь не выбран");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Создание шаблонов
        /// </summary>
        private async void buttonGenerate_Click(object sender, EventArgs e)
        {
            //Создаем файлы .            
            try
            {

                labelLoading.Visible = true; // лейбл состояния стал виден
                labelLoading.Text = "Загрузка..."; // изменение лейбла состояния
                int lastRow = TotalSize(_Excel.worksheetWorkPlanPlan)[0]; // Найти последний столбик листа, Excel файла
                int lastColumn = TotalSize(_Excel.worksheetWorkPlanPlan)[1]; // Найти последнюю строку листа, Excel файла
                MaxValueOfProgressBar(_Excel.worksheetWorkPlanPlan, lastRow, lastColumn); // Найти максимум прогресс бара
                await Task.Run(() =>
                {
                    Dictionary<string, string> dicTitle = PrepareDataFromSheetTitle(_Excel.worksheetWorkPlanTitlePage);
                    for (int i = 6; i <= lastRow; i++) // цикл проходящий по всем строкам
                    {
                        if (IsDiscipline(i, lastColumn))
                        {
                            Dictionary<string, string> dicPlan = PrepareDataFromSheetPlan(_Excel.worksheetWorkPlanPlan, i, lastColumn, dicTitle);
                            WriteInFile();
                            progressBar1.Value++;
                        }
                    }
                });
                labelLoading.Text = "Загрузка завершена";
                MessageBox.Show("Загрузка завершена");
                Reset();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Возвращает true если это дисциплина, иначе возвращает falce
        /// </summary>
        /// <param name="index"> номер строки </param>
        public bool IsDiscipline(int index, int lastColumn) =>
             _Excel.worksheetWorkPlanPlan.Cells[lastColumn - 1][index].Value != null || _Excel.worksheetWorkPlanPlan.Cells[10][index].Value != null;

        public void MaxValueOfProgressBar(Excel.Worksheet worksheet, int lastRow, int lastColumn)
        {
            int maxValueOfProgressBar = 0;
            for (int i = 6; i <= lastRow; i++)
                if (IsDiscipline(i, lastColumn))
                    maxValueOfProgressBar++;
            progressBar1.Maximum = maxValueOfProgressBar;
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
