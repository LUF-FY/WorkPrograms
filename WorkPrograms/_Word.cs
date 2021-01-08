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

        public void FillPattern(Dictionary<string, string> competenciesDic, Dictionary<string, string> dicTitle, Dictionary<string, string> dicPlan)
        {
            DocX document = DocX.Load("WordPattern.docx");
            //for (int i = 0; i < replaceableStrings.Count(); i++)
            //{
            //    string s = "$" + namesOfReplaceableStrings[i] + "$";
            //    string s2 = replaceableStrings[i];
            
            //    else if (namesOfReplaceableStrings[i] == "studyProgram")
            //    {
            //        SetStudyProgramTables(document, replaceableStrings[i]);
            //    }
            //    document.ReplaceText(s, s2);
            //}
            //if (!isInteractiveWatch)
            //{
            //    DeleteTable(3, document);
            //}
            FillDic(dicTitle, document);
            FillDic(dicPlan, document);
            CreateTable(competenciesDic, document);
            document.SaveAs(path);
        }

        private void SetStudyProgramTables(DocX document, string replaceableString)
        {
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
            DeleteTable(6, document);
            DeleteTable(5, document);
        }

        private void FillDic(Dictionary<string, string> dic, DocX document)
        {
            foreach (var el in dic)
            {
                if (el.Key == "$creditUnits$")
                {
                    document.ReplaceText(el.Key, ChangeDeclination(Convert.ToInt32(el.Value)));
                }
                else if (el.Key == "$studyProgram$")
                {
                    SetStudyProgramTables(document, el.Value);
                }
                else if (el.Key != "")
                    document.ReplaceText(el.Key, el.Value);
            }
        }

        private void DeleteTable(int number, DocX document)
        {
            var delTable = document.Tables[number];
            delTable.Remove();
        }

        private void CreateTable(Dictionary<string, string> competenciesDic, DocX document)
        {
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

            string s = $"{creditUnits} зачётных единиц.";
            if (creditUnits % 10 == 1) s = $"{creditUnits} зачётная единица.";
            if (creditUnits % 10 >= 2 && creditUnits % 10 <= 4) s = $"{creditUnits} зачётные единицы.";
            if (creditUnits % 100 >= 11 & creditUnits % 100 <= 20) s = $"{creditUnits} зачётных единиц.";
            return s;
        }
    }
}
