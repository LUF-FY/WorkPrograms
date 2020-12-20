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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void buttonOpenExcel_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult res = openFileDialogSelectFile.ShowDialog();
                if (res == DialogResult.OK)
                {
                    
                    //SelectFile.SelectExcelWorkPlanFile(openFileDialogSelectFile, labelNameOfWorkPlanFile);                    
                }
                else
                    throw new Exception("Файл не выбран");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
