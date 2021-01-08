
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WorkPrograms));
            this.openFileDialogSelectFile = new System.Windows.Forms.OpenFileDialog();
            this.buttonOpenExcel = new System.Windows.Forms.Button();
            this.labelNameOfWorkPlanFile = new System.Windows.Forms.Label();
            this.buttonGenerate = new System.Windows.Forms.Button();
            this.labelLoading = new System.Windows.Forms.Label();
            this.buttonOpenFolder = new System.Windows.Forms.Button();
            this.labelNameOfFolder = new System.Windows.Forms.Label();
            this.folderBrowserDialogChooseFolder = new System.Windows.Forms.FolderBrowserDialog();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.errorProvider1 = new System.Windows.Forms.ErrorProvider(this.components);
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.labelLastFile = new System.Windows.Forms.Label();
            this.labelNameOfLastFile = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).BeginInit();
            this.SuspendLayout();
            // 
            // openFileDialogSelectFile
            // 
            this.openFileDialogSelectFile.Filter = "Excel|*.xls|Excel|*.xlsx";
            // 
            // buttonOpenExcel
            // 
            this.buttonOpenExcel.BackColor = System.Drawing.Color.Transparent;
            this.buttonOpenExcel.Font = new System.Drawing.Font("Arial", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonOpenExcel.Location = new System.Drawing.Point(23, 40);
            this.buttonOpenExcel.Name = "buttonOpenExcel";
            this.buttonOpenExcel.Size = new System.Drawing.Size(140, 30);
            this.buttonOpenExcel.TabIndex = 0;
            this.buttonOpenExcel.Text = "Открыть";
            this.buttonOpenExcel.UseVisualStyleBackColor = false;
            this.buttonOpenExcel.Click += new System.EventHandler(this.buttonOpenExcel_Click);
            // 
            // labelNameOfWorkPlanFile
            // 
            this.labelNameOfWorkPlanFile.AutoSize = true;
            this.labelNameOfWorkPlanFile.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelNameOfWorkPlanFile.Location = new System.Drawing.Point(169, 49);
            this.labelNameOfWorkPlanFile.Name = "labelNameOfWorkPlanFile";
            this.labelNameOfWorkPlanFile.Size = new System.Drawing.Size(100, 15);
            this.labelNameOfWorkPlanFile.TabIndex = 1;
            this.labelNameOfWorkPlanFile.Text = "Файл не выбран";
            // 
            // buttonGenerate
            // 
            this.buttonGenerate.Enabled = false;
            this.buttonGenerate.Font = new System.Drawing.Font("Arial", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonGenerate.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.errorProvider1.SetIconAlignment(this.buttonGenerate, System.Windows.Forms.ErrorIconAlignment.TopLeft);
            this.buttonGenerate.Location = new System.Drawing.Point(23, 198);
            this.buttonGenerate.Name = "buttonGenerate";
            this.buttonGenerate.Size = new System.Drawing.Size(140, 31);
            this.buttonGenerate.TabIndex = 2;
            this.buttonGenerate.Text = "Сформировать";
            this.buttonGenerate.UseVisualStyleBackColor = true;
            this.buttonGenerate.Click += new System.EventHandler(this.buttonGenerate_Click);
            // 
            // labelLoading
            // 
            this.labelLoading.AutoSize = true;
            this.labelLoading.BackColor = System.Drawing.Color.Transparent;
            this.labelLoading.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelLoading.ForeColor = System.Drawing.SystemColors.ControlText;
            this.labelLoading.Location = new System.Drawing.Point(169, 180);
            this.labelLoading.Name = "labelLoading";
            this.labelLoading.Size = new System.Drawing.Size(70, 16);
            this.labelLoading.TabIndex = 3;
            this.labelLoading.Text = "Ожидание";
            this.labelLoading.Visible = false;
            // 
            // buttonOpenFolder
            // 
            this.buttonOpenFolder.BackColor = System.Drawing.Color.Transparent;
            this.buttonOpenFolder.Enabled = false;
            this.buttonOpenFolder.Font = new System.Drawing.Font("Arial", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonOpenFolder.Location = new System.Drawing.Point(23, 119);
            this.buttonOpenFolder.Margin = new System.Windows.Forms.Padding(2);
            this.buttonOpenFolder.Name = "buttonOpenFolder";
            this.buttonOpenFolder.Size = new System.Drawing.Size(140, 30);
            this.buttonOpenFolder.TabIndex = 4;
            this.buttonOpenFolder.Text = "Выбрать";
            this.buttonOpenFolder.UseVisualStyleBackColor = false;
            this.buttonOpenFolder.Click += new System.EventHandler(this.buttonOpenFolder_Click);
            // 
            // labelNameOfFolder
            // 
            this.labelNameOfFolder.AutoSize = true;
            this.labelNameOfFolder.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelNameOfFolder.Location = new System.Drawing.Point(167, 128);
            this.labelNameOfFolder.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.labelNameOfFolder.Name = "labelNameOfFolder";
            this.labelNameOfFolder.Size = new System.Drawing.Size(113, 15);
            this.labelNameOfFolder.TabIndex = 5;
            this.labelNameOfFolder.Text = "Папка не выбрана";
            // 
            // progressBar1
            // 
            this.progressBar1.BackColor = System.Drawing.SystemColors.Control;
            this.progressBar1.Location = new System.Drawing.Point(169, 199);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(401, 30);
            this.progressBar1.TabIndex = 6;
            // 
            // errorProvider1
            // 
            this.errorProvider1.ContainerControl = this;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(20, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(247, 17);
            this.label1.TabIndex = 7;
            this.label1.Text = "Выберите файл с учебным планом";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arial", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(20, 100);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(316, 17);
            this.label2.TabIndex = 8;
            this.label2.Text = "Выберите папку создания рабочих программ";
            // 
            // labelLastFile
            // 
            this.labelLastFile.AutoSize = true;
            this.labelLastFile.Font = new System.Drawing.Font("Arial", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelLastFile.Location = new System.Drawing.Point(434, 244);
            this.labelLastFile.Name = "labelLastFile";
            this.labelLastFile.Size = new System.Drawing.Size(226, 17);
            this.labelLastFile.TabIndex = 9;
            this.labelLastFile.Text = "Последний обработанный файл\r\n";
            // 
            // labelNameOfLastFile
            // 
            this.labelNameOfLastFile.AutoSize = true;
            this.labelNameOfLastFile.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelNameOfLastFile.Location = new System.Drawing.Point(434, 277);
            this.labelNameOfLastFile.Name = "labelNameOfLastFile";
            this.labelNameOfLastFile.Size = new System.Drawing.Size(118, 15);
            this.labelNameOfLastFile.TabIndex = 10;
            this.labelNameOfLastFile.Text = "Файл не обработан";
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Arial", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button1.Location = new System.Drawing.Point(23, 261);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(140, 31);
            this.button1.TabIndex = 11;
            this.button1.Text = "Прекратить";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // WorkPrograms
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(683, 311);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.labelNameOfLastFile);
            this.Controls.Add(this.labelLastFile);
            this.Controls.Add(this.labelLoading);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.labelNameOfFolder);
            this.Controls.Add(this.buttonOpenFolder);
            this.Controls.Add(this.buttonGenerate);
            this.Controls.Add(this.labelNameOfWorkPlanFile);
            this.Controls.Add(this.buttonOpenExcel);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "WorkPrograms";
            this.Text = "WorkPrograms";
            this.TransparencyKey = System.Drawing.Color.Transparent;
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialogSelectFile;
        private System.Windows.Forms.Button buttonOpenExcel;
        private System.Windows.Forms.Label labelNameOfWorkPlanFile;
        private System.Windows.Forms.Button buttonGenerate;
        private System.Windows.Forms.Button buttonOpenFolder;
        private System.Windows.Forms.Label labelNameOfFolder;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialogChooseFolder;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.ErrorProvider errorProvider1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        public System.Windows.Forms.Label labelLoading;
        private System.Windows.Forms.Label labelNameOfLastFile;
        private System.Windows.Forms.Label labelLastFile;
        private System.Windows.Forms.Button button1;
    }
}

