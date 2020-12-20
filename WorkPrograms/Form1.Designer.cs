
namespace WorkPrograms
{
    partial class WorkPrograms
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.openFileDialogSelectFile = new System.Windows.Forms.OpenFileDialog();
            this.buttonOpenExcel = new System.Windows.Forms.Button();
            this.labelNameOfWorkPlanFile = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // openFileDialogSelectFile
            // 
            this.openFileDialogSelectFile.FileName = "openFileDialog1";
            // 
            // buttonOpenExcel
            // 
            this.buttonOpenExcel.Location = new System.Drawing.Point(27, 36);
            this.buttonOpenExcel.Name = "buttonOpenExcel";
            this.buttonOpenExcel.Size = new System.Drawing.Size(110, 26);
            this.buttonOpenExcel.TabIndex = 0;
            this.buttonOpenExcel.Text = "Открыть файл";
            this.buttonOpenExcel.UseVisualStyleBackColor = true;
            this.buttonOpenExcel.Click += new System.EventHandler(this.buttonOpenExcel_Click);
            // 
            // labelNameOfWorkPlanFile
            // 
            this.labelNameOfWorkPlanFile.AutoSize = true;
            this.labelNameOfWorkPlanFile.Location = new System.Drawing.Point(155, 43);
            this.labelNameOfWorkPlanFile.Name = "labelNameOfWorkPlanFile";
            this.labelNameOfWorkPlanFile.Size = new System.Drawing.Size(92, 13);
            this.labelNameOfWorkPlanFile.TabIndex = 1;
            this.labelNameOfWorkPlanFile.Text = "Файл не выбран";
            // 
            // WorkPrograms
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(584, 261);
            this.Controls.Add(this.labelNameOfWorkPlanFile);
            this.Controls.Add(this.buttonOpenExcel);
            this.Name = "WorkPrograms";
            this.Text = "WorkPrograms";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialogSelectFile;
        private System.Windows.Forms.Button buttonOpenExcel;
        private System.Windows.Forms.Label labelNameOfWorkPlanFile;
    }
}

