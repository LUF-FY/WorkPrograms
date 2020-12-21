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

        public void FillPattern()
        {
            DocX document = DocX.Load(path);
            string[] replaceableStrings = new string[]
            {
                WorkPrograms.subjectName, WorkPrograms.direction, WorkPrograms.profile,
                WorkPrograms.standard, WorkPrograms.protocol, WorkPrograms.chair,  
                WorkPrograms.creditUnits.ToString(), WorkPrograms.studyHours.ToString(),
                WorkPrograms.courses, WorkPrograms.semesters, WorkPrograms.sumIndependentWork.ToString(),
                WorkPrograms.typesOfLessons, WorkPrograms.test
            };
            foreach(var el in replaceableStrings)
            {
                string s = "$" + el + "$";
                document.ReplaceText(s, el);
            }
            foreach(var el in WorkPrograms.semesterData)
            {
                if (el.Key != "")
                    document.ReplaceText(el.Key, el.Value);
            }
            document.Save();
        }
    }
}
