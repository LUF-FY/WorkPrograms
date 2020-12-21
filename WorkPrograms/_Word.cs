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

        public void FillPattern(string competencies)
        {
            DocX document = DocX.Load("WordPattern.docx");
            string[] replaceableStrings = new string[]
            {
                WorkPrograms.subjectName, WorkPrograms.direction, WorkPrograms.profile,
                WorkPrograms.standard, WorkPrograms.protocol, WorkPrograms.chair,  
                WorkPrograms.creditUnits.ToString(), WorkPrograms.studyHours,
                WorkPrograms.courses, WorkPrograms.semesters, WorkPrograms.sumIndependentWork.ToString(),
                WorkPrograms.typesOfLessons, WorkPrograms.test, WorkPrograms.consulting, WorkPrograms.courseWork,
                competencies, WorkPrograms.edForm
            };
            string[] namesOfReplaceableStrings = new string[]
            {
                nameof(WorkPrograms.subjectName), nameof(WorkPrograms.direction), nameof(WorkPrograms.profile),
                nameof(WorkPrograms.standard), nameof(WorkPrograms.protocol), nameof(WorkPrograms.chair),
                nameof(WorkPrograms.creditUnits), nameof(WorkPrograms.studyHours),
                nameof(WorkPrograms.courses), nameof(WorkPrograms.semesters), nameof(WorkPrograms.sumIndependentWork),
                nameof(WorkPrograms.typesOfLessons), nameof(WorkPrograms.test), nameof(WorkPrograms.consulting), nameof(WorkPrograms.courseWork),
                nameof(competencies), nameof(WorkPrograms.edForm)
            };
            for (int i = 0; i < replaceableStrings.Count(); i++)
            {
                string s = "";
                if (replaceableStrings[i] == WorkPrograms.creditUnits.ToString())
                {
                    s = "$" + namesOfReplaceableStrings[i] + "$";
                    string s2 = ChangeDeclination(WorkPrograms.creditUnits);
                    document.ReplaceText(s, s2);
                }
                else
                    s = "$" + namesOfReplaceableStrings[i] + "$";
                document.ReplaceText(s, replaceableStrings[i]);
            }
            foreach(var el in WorkPrograms.semesterData)
            {
                if (el.Key != "")
                    document.ReplaceText(el.Key, el.Value);
            }
            document.SaveAs(path);
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
