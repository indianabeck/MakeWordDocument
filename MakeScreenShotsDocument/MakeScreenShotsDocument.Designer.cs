namespace MakeScreenShotsDocument
{
    partial class MakeScreenShotsDocument
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MakeScreenShotsDocument));
            this.LeftBrowser = new System.Windows.Forms.TextBox();
            this.RightBrowser = new System.Windows.Forms.TextBox();
            this.LeftScreenShotsPath = new System.Windows.Forms.TextBox();
            this.RightScreenShotsPath = new System.Windows.Forms.TextBox();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.LeftScreenShotsBrowse = new System.Windows.Forms.Button();
            this.RightScreenShotsBrowse = new System.Windows.Forms.Button();
            this.ExcelFile = new System.Windows.Forms.TextBox();
            this.ExcelFileBrowser = new System.Windows.Forms.Button();
            this.HeaderPage = new System.Windows.Forms.TextBox();
            this.HeaderPageBrowser = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.CreateDocument = new System.Windows.Forms.Button();
            this.LeftBrowserLabel = new System.Windows.Forms.Label();
            this.LeftScreenShotPathLabel = new System.Windows.Forms.Label();
            this.RightBrowserLabel = new System.Windows.Forms.Label();
            this.ExcelFileLabel = new System.Windows.Forms.Label();
            this.HeaderFileLabel = new System.Windows.Forms.Label();
            this.RightScreenShotPathLabel = new System.Windows.Forms.Label();
            this.Help = new System.Windows.Forms.Button();
            this.CreateExcelDocument = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.Url = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.OverrideExtension = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // LeftBrowser
            // 
            this.LeftBrowser.Location = new System.Drawing.Point(29, 26);
            this.LeftBrowser.Name = "LeftBrowser";
            this.LeftBrowser.Size = new System.Drawing.Size(272, 20);
            this.LeftBrowser.TabIndex = 0;
            this.LeftBrowser.TextChanged += new System.EventHandler(this.LeftBrowser_TextChanged);
            // 
            // RightBrowser
            // 
            this.RightBrowser.Location = new System.Drawing.Point(29, 108);
            this.RightBrowser.Name = "RightBrowser";
            this.RightBrowser.Size = new System.Drawing.Size(335, 20);
            this.RightBrowser.TabIndex = 2;
            this.RightBrowser.TextChanged += new System.EventHandler(this.RightBrowser_TextChanged);
            // 
            // LeftScreenShotsPath
            // 
            this.LeftScreenShotsPath.Location = new System.Drawing.Point(29, 65);
            this.LeftScreenShotsPath.Name = "LeftScreenShotsPath";
            this.LeftScreenShotsPath.Size = new System.Drawing.Size(503, 20);
            this.LeftScreenShotsPath.TabIndex = 1;
            this.LeftScreenShotsPath.TextChanged += new System.EventHandler(this.LeftScreenShotsPath_TextChanged);
            // 
            // RightScreenShotsPath
            // 
            this.RightScreenShotsPath.Location = new System.Drawing.Point(29, 149);
            this.RightScreenShotsPath.Name = "RightScreenShotsPath";
            this.RightScreenShotsPath.Size = new System.Drawing.Size(503, 20);
            this.RightScreenShotsPath.TabIndex = 3;
            this.RightScreenShotsPath.TextChanged += new System.EventHandler(this.RightScreenShotsPath_TextChanged);
            // 
            // LeftScreenShotsBrowse
            // 
            this.LeftScreenShotsBrowse.Location = new System.Drawing.Point(550, 65);
            this.LeftScreenShotsBrowse.Name = "LeftScreenShotsBrowse";
            this.LeftScreenShotsBrowse.Size = new System.Drawing.Size(40, 20);
            this.LeftScreenShotsBrowse.TabIndex = 0;
            this.LeftScreenShotsBrowse.TabStop = false;
            this.LeftScreenShotsBrowse.Text = "...";
            this.LeftScreenShotsBrowse.UseVisualStyleBackColor = true;
            this.LeftScreenShotsBrowse.Click += new System.EventHandler(this.LeftScreenShotsBrowse_Click);
            // 
            // RightScreenShotsBrowse
            // 
            this.RightScreenShotsBrowse.Location = new System.Drawing.Point(550, 148);
            this.RightScreenShotsBrowse.Name = "RightScreenShotsBrowse";
            this.RightScreenShotsBrowse.Size = new System.Drawing.Size(40, 20);
            this.RightScreenShotsBrowse.TabIndex = 0;
            this.RightScreenShotsBrowse.TabStop = false;
            this.RightScreenShotsBrowse.Text = "...";
            this.RightScreenShotsBrowse.UseVisualStyleBackColor = true;
            this.RightScreenShotsBrowse.Click += new System.EventHandler(this.RightScreenShotsBrowse_Click);
            // 
            // ExcelFile
            // 
            this.ExcelFile.Location = new System.Drawing.Point(29, 188);
            this.ExcelFile.Name = "ExcelFile";
            this.ExcelFile.Size = new System.Drawing.Size(503, 20);
            this.ExcelFile.TabIndex = 4;
            this.ExcelFile.TextChanged += new System.EventHandler(this.ExcelFile_TextChanged);
            // 
            // ExcelFileBrowser
            // 
            this.ExcelFileBrowser.Location = new System.Drawing.Point(550, 188);
            this.ExcelFileBrowser.Name = "ExcelFileBrowser";
            this.ExcelFileBrowser.Size = new System.Drawing.Size(40, 20);
            this.ExcelFileBrowser.TabIndex = 0;
            this.ExcelFileBrowser.TabStop = false;
            this.ExcelFileBrowser.Text = "...";
            this.ExcelFileBrowser.UseVisualStyleBackColor = true;
            this.ExcelFileBrowser.Click += new System.EventHandler(this.ExcelFileBrowser_Click);
            // 
            // HeaderPage
            // 
            this.HeaderPage.Location = new System.Drawing.Point(29, 227);
            this.HeaderPage.Name = "HeaderPage";
            this.HeaderPage.Size = new System.Drawing.Size(503, 20);
            this.HeaderPage.TabIndex = 5;
            this.HeaderPage.TextChanged += new System.EventHandler(this.HeaderPage_TextChanged);
            // 
            // HeaderPageBrowser
            // 
            this.HeaderPageBrowser.Location = new System.Drawing.Point(550, 227);
            this.HeaderPageBrowser.Name = "HeaderPageBrowser";
            this.HeaderPageBrowser.Size = new System.Drawing.Size(40, 20);
            this.HeaderPageBrowser.TabIndex = 0;
            this.HeaderPageBrowser.TabStop = false;
            this.HeaderPageBrowser.Text = "...";
            this.HeaderPageBrowser.UseVisualStyleBackColor = true;
            this.HeaderPageBrowser.Click += new System.EventHandler(this.HeaderPageBrowser_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // CreateDocument
            // 
            this.CreateDocument.Location = new System.Drawing.Point(29, 316);
            this.CreateDocument.Name = "CreateDocument";
            this.CreateDocument.Size = new System.Drawing.Size(126, 23);
            this.CreateDocument.TabIndex = 7;
            this.CreateDocument.Text = "Create Document";
            this.CreateDocument.UseVisualStyleBackColor = true;
            this.CreateDocument.Click += new System.EventHandler(this.CreateDocument_Click);
            // 
            // LeftBrowserLabel
            // 
            this.LeftBrowserLabel.Location = new System.Drawing.Point(26, 10);
            this.LeftBrowserLabel.Name = "LeftBrowserLabel";
            this.LeftBrowserLabel.Size = new System.Drawing.Size(506, 13);
            this.LeftBrowserLabel.TabIndex = 7;
            this.LeftBrowserLabel.Text = "Left Browser";
            // 
            // LeftScreenShotPathLabel
            // 
            this.LeftScreenShotPathLabel.Location = new System.Drawing.Point(26, 49);
            this.LeftScreenShotPathLabel.Name = "LeftScreenShotPathLabel";
            this.LeftScreenShotPathLabel.Size = new System.Drawing.Size(506, 13);
            this.LeftScreenShotPathLabel.TabIndex = 8;
            this.LeftScreenShotPathLabel.Text = "Left Screen Shots Path";
            // 
            // RightBrowserLabel
            // 
            this.RightBrowserLabel.Location = new System.Drawing.Point(26, 92);
            this.RightBrowserLabel.Name = "RightBrowserLabel";
            this.RightBrowserLabel.Size = new System.Drawing.Size(506, 13);
            this.RightBrowserLabel.TabIndex = 9;
            this.RightBrowserLabel.Text = "Right Browser (leave blank if you don\'t want to populate the right side)";
            // 
            // ExcelFileLabel
            // 
            this.ExcelFileLabel.Location = new System.Drawing.Point(26, 172);
            this.ExcelFileLabel.Name = "ExcelFileLabel";
            this.ExcelFileLabel.Size = new System.Drawing.Size(506, 13);
            this.ExcelFileLabel.TabIndex = 10;
            this.ExcelFileLabel.Text = "Excel File (Full Path and Name)";
            // 
            // HeaderFileLabel
            // 
            this.HeaderFileLabel.Location = new System.Drawing.Point(26, 207);
            this.HeaderFileLabel.Name = "HeaderFileLabel";
            this.HeaderFileLabel.Size = new System.Drawing.Size(506, 13);
            this.HeaderFileLabel.TabIndex = 11;
            this.HeaderFileLabel.Text = "Header Doc File (Full Path and Name - Leave blank for an empty Header Page)";
            // 
            // RightScreenShotPathLabel
            // 
            this.RightScreenShotPathLabel.Location = new System.Drawing.Point(26, 133);
            this.RightScreenShotPathLabel.Name = "RightScreenShotPathLabel";
            this.RightScreenShotPathLabel.Size = new System.Drawing.Size(506, 13);
            this.RightScreenShotPathLabel.TabIndex = 12;
            this.RightScreenShotPathLabel.Text = "Right Screen Shots Path";
            // 
            // Help
            // 
            this.Help.Location = new System.Drawing.Point(208, 316);
            this.Help.Name = "Help";
            this.Help.Size = new System.Drawing.Size(126, 23);
            this.Help.TabIndex = 13;
            this.Help.Text = "Help";
            this.Help.UseVisualStyleBackColor = true;
            this.Help.Click += new System.EventHandler(this.Help_Click);
            // 
            // CreateExcelDocument
            // 
            this.CreateExcelDocument.Location = new System.Drawing.Point(370, 316);
            this.CreateExcelDocument.Name = "CreateExcelDocument";
            this.CreateExcelDocument.Size = new System.Drawing.Size(126, 23);
            this.CreateExcelDocument.TabIndex = 14;
            this.CreateExcelDocument.Text = "Create Excel Doc";
            this.CreateExcelDocument.UseVisualStyleBackColor = true;
            this.CreateExcelDocument.Click += new System.EventHandler(this.CreateExcelDocument_Click);
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(26, 250);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(506, 13);
            this.label1.TabIndex = 16;
            this.label1.Text = "Url for Page Titles";
            // 
            // Url
            // 
            this.Url.Location = new System.Drawing.Point(29, 266);
            this.Url.Name = "Url";
            this.Url.Size = new System.Drawing.Size(503, 20);
            this.Url.TabIndex = 6;
            this.Url.LostFocus += new System.EventHandler(this.Url_LostFocus);
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(367, 92);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(165, 13);
            this.label2.TabIndex = 18;
            this.label2.Text = "Override Extension";
            // 
            // OverrideExtension
            // 
            this.OverrideExtension.Location = new System.Drawing.Point(370, 108);
            this.OverrideExtension.Name = "OverrideExtension";
            this.OverrideExtension.Size = new System.Drawing.Size(162, 20);
            this.OverrideExtension.TabIndex = 17;
            this.OverrideExtension.TabStop = false;
            this.OverrideExtension.TextChanged += new System.EventHandler(this.OverrideExtension_TextChanged);
            // 
            // MakeScreenShotsDocument
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(615, 351);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.OverrideExtension);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Url);
            this.Controls.Add(this.CreateExcelDocument);
            this.Controls.Add(this.Help);
            this.Controls.Add(this.RightScreenShotPathLabel);
            this.Controls.Add(this.HeaderFileLabel);
            this.Controls.Add(this.ExcelFileLabel);
            this.Controls.Add(this.RightBrowserLabel);
            this.Controls.Add(this.LeftScreenShotPathLabel);
            this.Controls.Add(this.LeftBrowserLabel);
            this.Controls.Add(this.CreateDocument);
            this.Controls.Add(this.HeaderPageBrowser);
            this.Controls.Add(this.HeaderPage);
            this.Controls.Add(this.ExcelFileBrowser);
            this.Controls.Add(this.ExcelFile);
            this.Controls.Add(this.RightScreenShotsBrowse);
            this.Controls.Add(this.LeftScreenShotsBrowse);
            this.Controls.Add(this.RightScreenShotsPath);
            this.Controls.Add(this.LeftScreenShotsPath);
            this.Controls.Add(this.RightBrowser);
            this.Controls.Add(this.LeftBrowser);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimumSize = new System.Drawing.Size(631, 390);
            this.Name = "MakeScreenShotsDocument";
            this.Text = "Make ScreenShots Document";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox LeftBrowser;
        private System.Windows.Forms.TextBox RightBrowser;
        private System.Windows.Forms.TextBox LeftScreenShotsPath;
        private System.Windows.Forms.TextBox RightScreenShotsPath;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button LeftScreenShotsBrowse;
        private System.Windows.Forms.Button RightScreenShotsBrowse;
        private System.Windows.Forms.TextBox ExcelFile;
        private System.Windows.Forms.Button ExcelFileBrowser;
        private System.Windows.Forms.TextBox HeaderPage;
        private System.Windows.Forms.Button HeaderPageBrowser;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button CreateDocument;
        private System.Windows.Forms.Label LeftBrowserLabel;
        private System.Windows.Forms.Label LeftScreenShotPathLabel;
        private System.Windows.Forms.Label RightBrowserLabel;
        private System.Windows.Forms.Label ExcelFileLabel;
        private System.Windows.Forms.Label HeaderFileLabel;
        private System.Windows.Forms.Label RightScreenShotPathLabel;
        private System.Windows.Forms.Button Help;
        private System.Windows.Forms.Button CreateExcelDocument;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox Url;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox OverrideExtension;
    }
}

