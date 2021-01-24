using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;

namespace WorkPrograms
{
    class _Word
    {
        public string path;

        public void FillPattern(Dictionary<string, string> competenciesDic, Dictionary<string, string> dicTitle, 
            Dictionary<string, string> dicPlan)
        {
            // Заполнение шаблона Word.
            DocX document = DocX.Load("WordPattern.docx");
            ReplaceTestOrExam(dicPlan, document);
            ReplaceFromDic(dicTitle, document);
            ReplaceFromDic(dicPlan, document);
            CreateTable(competenciesDic, document);
            document.SaveAs(path);
        }

        private void SetStudyProgramTables(DocX document, string replaceableString)
        {
            // Изменение шаблона в зависимости от программы обучения(бакалавриат/магистратура/аспирантура).
            if (replaceableString != "бакалавриата")
            {
                document.ReplaceTextWithObject("$table5$", document.Tables[6]);
                document.ReplaceText("Таблица 8.1", "");
                document.ReplaceText("Критерии оценивания представлены в таблице 8.1", "");
                document.ReplaceText("Методика формирования результирующей оценки", "");
                DeleteTable(4, document);
                if (replaceableString == "магистратуры")
                {
                    document.ReplaceText("$школьного курса$", "бакалавриата");
                }
                else if (replaceableString == "аспирантуры")
                {
                    document.ReplaceText("$школьного курса$", "бакалавриата, магистратуры или специалитета");
                }
            }
            else
            {
                document.ReplaceTextWithObject("$table5$", document.Tables[5]);
                document.ReplaceText("$школьного курса$", "школьного курса");
            }
            DeleteTable(document.Tables.Count - 2, document);
            DeleteTable(document.Tables.Count - 1, document);
        }

        private void ReplaceTestOrExam(Dictionary<string, string> dicPlan, DocX document)
        {
            // Выбор зачет/экзамен.
            if (dicPlan["$exam$"] == "-")
                document.ReplaceText("$testOrExam$", "зачёту");
            else if (dicPlan["$test$"] == "-")
                document.ReplaceText("$testOrExam$", "экзамену");
            else
                document.ReplaceText("$testOrExam$", "зачёту/экзамену");
        }

        private void ReplaceFromDic(Dictionary<string, string> dic, DocX document)
        {
            // Замена кодовых слов на нужные значения из словаря.
            foreach (var el in dic)
            {
                if (el.Key == "$creditUnits$")
                {
                    document.ReplaceText(el.Key, ChangeDeclination(Convert.ToInt32(el.Value)));
                }
                else if (el.Key == "$studyProgram$")
                {
                    document.ReplaceText(el.Key, el.Value);
                    SetStudyProgramTables(document, el.Value);
                }
                else if (el.Key == "$interactiveWatch$" && el.Value == "")
                {
                    DeleteTable(3, document);
                }
                else if (el.Key == "$profile$" && el.Value == "")
                {
                    document.ReplaceText("$profile$, ", "");
                    document.ReplaceText(el.Key, el.Value);
                }
                else if (el.Key != "")
                {
                    document.ReplaceText(el.Key, el.Value);
                }
            }
        }

        private void DeleteTable(int number, DocX document)
        {
            // Удаление таблицы из шаблона.
            var delTable = document.Tables[number];
            delTable.Remove();
        }

        private void CreateTable(Dictionary<string, string> competenciesDic, DocX document)
        {
            // Создание таблицы с компетенциями.
            var compTable = document.Tables[1];
            var compList = WorkPrograms.subjectCompetencies.Split(';', ' ').ToList();
            foreach (var item in compList)
            {
                if (!string.IsNullOrEmpty(item))
                {
                    if (competenciesDic.ContainsKey(item))
                    {
                        var row = compTable.InsertRow();
                        row.Cells[0].Paragraphs[0].Append(item);
                        row.Cells[1].Paragraphs[0].Append(competenciesDic[item]);
                        for (int i = 2; i < compTable.ColumnCount; i++)
                        {
                            row.Cells[i].Paragraphs[0].Append("Вставка").Highlight(Xceed.Document.NET.Highlight.cyan);
                        }
                    }
                }
            }
        }

        private string ChangeDeclination(int creditUnits)
        {
            // Склонение зачетных ед.
            string s = $"{creditUnits} зачётных единиц.";
            if (creditUnits % 10 == 1) s = $"{creditUnits} зачётная единица.";
            if (creditUnits % 10 >= 2 && creditUnits % 10 <= 4) s = $"{creditUnits} зачётные единицы.";
            if (creditUnits % 100 >= 11 & creditUnits % 100 <= 20) s = $"{creditUnits} зачётных единиц.";
            return s;
        }
    }
}
