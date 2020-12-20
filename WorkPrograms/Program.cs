using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WorkPrograms
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new WorkPrograms());
            if (_Excel.xlApp != null)
            {
                _Excel.xlApp.Quit();
                _Excel.ClearExcel();
            }
        }
    }
}
