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

        public void FillPattern(Dictionary<string, string> competenciesDic, string[] replaceableStrings, 
            string[] namesOfReplaceableStrings, Dictionary<string, string> semesterData, bool isInteractiveWatch)
        {
            DocX document = DocX.Load("WordPattern.docx");
            for (int i = 0; i < replaceableStrings.Count(); i++)
            {
                string s = "$" + namesOfReplaceableStrings[i] + "$";
                string s2 = replaceableStrings[i];
                if (namesOfReplaceableStrings[i] == "creditUnits")
                {
                    s = "$" + namesOfReplaceableStrings[i] + "$";
                    s2 = ChangeDeclination(Convert.ToInt32(replaceableStrings[i]));
                }
                else if (namesOfReplaceableStrings[i] == "sumWorkshops")
                {
                    if (int.Parse(replaceableStrings[i]) == 0 && WorkPrograms.sumLaboratoryExercises != 0)
                    {
                        document.ReplaceText("$labOrPract$", "лаб");
                        s2 = WorkPrograms.sumLaboratoryExercises.ToString();
                    }
                    else
                        document.ReplaceText("$labOrPract$", "пр");
                }
                document.ReplaceText(s, s2);
            }

            foreach(var el in semesterData)
            {
                if (el.Key != "")
                    document.ReplaceText(el.Key, el.Value);
            }

            CreateTable(competenciesDic, document);

            if (!isInteractiveWatch)
            {
                Xceed.Document.NET.Table delTable = document.Tables[3];
                delTable.Remove();
            }
            document.SaveAs(path);
        }

        private void CreateTable(Dictionary<string, string> competenciesDic, DocX document)
        {
            Xceed.Document.NET.Table compTable = document.Tables[1];
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
