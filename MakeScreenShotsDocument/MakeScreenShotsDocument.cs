using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

using System.Net;
using System.Diagnostics;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

using System;
using System.IO;
using System.Collections;

namespace MakeScreenShotsDocument
{
    public partial class MakeScreenShotsDocument : Form
    {
        public MakeScreenShotsDocument()
        {
            InitializeComponent();
            // Need these in Initialized Components
            this.FormClosed += MakeScreenShotsDocument_FormClosed;                          //added
            this.Activated += MakeScreenShotsDocument_Activated;                            //added
            //this.Url.LostFocus += new System.EventHandler(this.Url_LostFocus);              //added
            LeftBrowser.Text = Properties.Settings.Default.LeftBrowser.ToString();
            LeftScreenShotsPath.Text = Properties.Settings.Default.LeftScreenShotsPath.ToString();
            RightBrowser.Text = Properties.Settings.Default.RightBrowser.ToString();
            RightScreenShotsPath.Text = Properties.Settings.Default.RightScreenShotsPath.ToString();
            ExcelFile.Text = Properties.Settings.Default.ExcelFile.ToString();
            HeaderPage.Text = Properties.Settings.Default.HeaderPage.ToString();
            Url.Text = Properties.Settings.Default.Url.ToString();
            OverrideExtension.Text = Properties.Settings.Default.OverrideExtension.ToString();
        }
        #region GetWebPageTitle
        public string GetWebPageTitle(string url)
        {
            string title;
            WebClient wc = new WebClient();
            try {
                string source = wc.DownloadString(url);
                title = Regex.Match(source, @"\<title\b[^>]*\>\s*(?<Title>[\s\S]*?)\</title\>", RegexOptions.IgnoreCase).Groups["Title"].Value;
                Encoding ascii = Encoding.ASCII;
                Encoding unicode = Encoding.Unicode;

                byte[] unicodeBytes = unicode.GetBytes(title);
                byte[] asciiBytes = Encoding.Convert(unicode, ascii, unicodeBytes);
                char[] asciiChars = new char[ascii.GetCharCount(asciiBytes, 0, asciiBytes.Length)];
                ascii.GetChars(asciiBytes, 0, asciiBytes.Length, asciiChars, 0);
                string asciiString = new string(asciiChars);
                title = asciiString;
                title = title.Replace("\r\n", "").Replace("\"", "").Replace("\t", "");
            }
            catch (Exception ex) {
                title = "weberror";
            }
            wc.Dispose();
            return title;
        }
        #endregion

        #region PopulateExcel
        private void PopulateExcel(string url)
        {
            //string url = "http://www.CialisMD.com/Pages/";
            string fullurl;
            string pageName;
            string lastTitle = String.Empty;
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Add();
            Excel.Sheets xlSheets = xlWorkbook.Worksheets;
            Excel.Worksheet xlSheet = (Excel.Worksheet)xlSheets[1];
            xlApp.Visible = true;
            Excel.Range xlRange = xlSheet.get_Range("A1", "A1");
            xlRange.Value= "Override";
            xlRange = xlSheet.get_Range("B1", "B1");
            xlRange.Value = "Image";
            xlRange = xlSheet.get_Range("C1", "C1");
            xlRange.Value = "Title";

            string lastPage = "";
            string PageTitle;
            int xlRow = 2;
            string xlCell = "A";
            string leftFilePath = this.LeftScreenShotsPath.Text.ToString();

            if (!leftFilePath.EndsWith(@"\")) {
                leftFilePath += @"\";
            }
            // Process the list of files found in the directory. 
            string[] fileEntries = Directory.GetFiles(leftFilePath);
            foreach (string fileName in fileEntries) {
                string slot = xlCell + xlRow.ToString();
                xlRange = xlSheet.get_Range(slot, slot);
                slot = "B" + xlRow.ToString();
                xlRange.Value = "30";
                xlRange = xlSheet.get_Range(slot, slot);
                pageName = Path.GetFileName(fileName);
                string[] words = pageName.ToString().Split('.');
                xlRange.Value = pageName;
                slot = "C" + xlRow.ToString();
                xlRange = xlSheet.get_Range(slot, slot);
                words = pageName.ToString().Split('_');
                if (words[0].Contains(".")) {
                    words = words[0].ToString().Split('.');
                }
                fullurl = url + words[0] + ".aspx";
                if (fullurl != lastPage) {
                    PageTitle = GetWebPageTitle(fullurl);
                    PageTitle = WebUtility.HtmlDecode(PageTitle);
                    lastTitle = PageTitle;
                }
                else {
                    PageTitle = lastTitle;
                }
                lastPage = fullurl;
                xlRange.Value = PageTitle;
                xlRow++;
            }

            xlSheet.Range["A1:C999"].EntireColumn.ColumnWidth = 135;
            xlSheet.Range["A1:C999"].Style.WrapText = false;
            xlSheet.Range["A1:C999"].Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            xlSheet.Range["A1:C999"].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
            xlSheet.Range["A1:C999"].EntireColumn.AutoFit();
            xlSheet.Range["A1:C999"].EntireRow.AutoFit();
            xlSheet.Range["A1:C999"].EntireColumn.RowHeight = 15;
            xlSheet.Application.ActiveWindow.SplitRow = 1;
            xlSheet.Application.ActiveWindow.FreezePanes = true;
            
            //xlWorkbook.Close();
            xlApp.Quit();
        }
        #endregion
        public void SelectAFolder(ref TextBox box)
        {
            if (box.Text.ToString() != "") {
                folderBrowserDialog1.SelectedPath = box.Text.ToString();
            }
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK) {
                string path = folderBrowserDialog1.SelectedPath;
                if(!path.EndsWith(@"\")) {
                    path = path + @"\";
                }
                box.Text = path;
            }
        }

        public void SelectAFile(ref TextBox box)
        {
            if (box.Name.ToString().ToLower().Contains("excel")) {
                openFileDialog1.Filter = "Excel Files|*xls;*.xlsx";
            }
            else if (box.Name.ToString().ToLower().Contains("header")) {
                openFileDialog1.Filter = "Word Files|*doc;*.docx";
            }
            else {
                openFileDialog1.Filter = "";
            }

            if (box.Text.ToString() != "") {
                openFileDialog1.FileName = box.Text.ToString();
            }
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK) {
                string path = openFileDialog1.FileName;
                box.Text = path;
            }
        }

        public void SaveSetting(TextBox box)
        {
            string boxText = box.Text.ToString();
            string boxName = box.Name.ToString();
            switch (boxName) {
                case "LeftBrowser":
                    Properties.Settings.Default.LeftBrowser = boxText;
                    break;
                case "LeftScreenShotsPath":
                    Properties.Settings.Default.LeftScreenShotsPath = boxText;
                    break;
                case "RightBrowser":
                    Properties.Settings.Default.RightBrowser = boxText;
                    break;
                case "RightScreenShotsPath":
                    Properties.Settings.Default.RightScreenShotsPath = boxText;
                    break;
                case "ExcelFile":
                    Properties.Settings.Default.ExcelFile = boxText;
                    break;
                case "HeaderPage":
                    Properties.Settings.Default.HeaderPage = boxText;
                    break;
                case "Url":
                    Properties.Settings.Default.Url = boxText;
                    break;
                case "OverrideExtension":
                    Properties.Settings.Default.OverrideExtension = boxText;
                    break;
                default:
                    break;
            }
            Properties.Settings.Default.Save();
        }

        private void LeftScreenShotsBrowse_Click(object sender, EventArgs e)
        {
            SelectAFolder(ref LeftScreenShotsPath);
            SaveSetting(LeftScreenShotsPath);
            LeftScreenShotsPath.Focus();
        }

        private void RightScreenShotsBrowse_Click(object sender, EventArgs e)
        {
            SelectAFolder(ref RightScreenShotsPath);
            SaveSetting(RightScreenShotsPath);
            RightScreenShotsPath.Focus();
        }

        private void ExcelFileBrowser_Click(object sender, EventArgs e)
        {
            SelectAFile(ref ExcelFile);
            SaveSetting(ExcelFile);
            ExcelFile.Focus();
        }

        private void HeaderPageBrowser_Click(object sender, EventArgs e)
        {
            SelectAFile(ref HeaderPage);
            SaveSetting(HeaderPage);
            HeaderPage.Focus();
        }

        private void LeftBrowser_TextChanged(object sender, EventArgs e)
        {
            SaveSetting(LeftBrowser);
        }
        private void LeftScreenShotsPath_TextChanged(object sender, EventArgs e)
        {
            SaveSetting(LeftScreenShotsPath);

        }
        private void RightBrowser_TextChanged(object sender, EventArgs e)
        {
            SaveSetting(RightBrowser);
        }
        private void RightScreenShotsPath_TextChanged(object sender, EventArgs e)
        {
            SaveSetting(RightScreenShotsPath);
        }
        private void ExcelFile_TextChanged(object sender, EventArgs e)
        {
            SaveSetting(ExcelFile);
        }
        private void HeaderPage_TextChanged(object sender, EventArgs e)
        {
            SaveSetting(HeaderPage);
        }
        private void OverrideExtension_TextChanged(object sender, System.EventArgs e)
        {
            SaveSetting(OverrideExtension);
        }
        //private void Url_TextChanged(object sender, System.EventArgs e)
        private void Url_LostFocus(object sender, System.EventArgs e)
        {
            if (!Url.Text.ToString().EndsWith(@"/")) {
                Url.Text = Url.Text.ToString() + @"/";
            }
            SaveSetting(Url);
        }
        private void MakeScreenShotsDocument_FormClosed(object sender, FormClosedEventArgs e)
        {
            // Need these in Initialized Components
            //this.FormClosed += MakeScreenShotsDocument_FormClosed;                          //added
            //this.Activated += MakeScreenShotsDocument_Activated;                            //added

            if (WindowState == FormWindowState.Maximized) {
                Properties.Settings.Default.Location = RestoreBounds.Location;
                Properties.Settings.Default.Size = RestoreBounds.Size;
                Properties.Settings.Default.Maximised = true;
                Properties.Settings.Default.Minimised = false;
            }
            else if (WindowState == FormWindowState.Normal) {
                Properties.Settings.Default.Location = Location;
                Properties.Settings.Default.Size = Size;
                Properties.Settings.Default.Maximised = false;
                Properties.Settings.Default.Minimised = false;
            }
            else {
                Properties.Settings.Default.Location = RestoreBounds.Location;
                Properties.Settings.Default.Size = RestoreBounds.Size;
                Properties.Settings.Default.Maximised = false;
                Properties.Settings.Default.Minimised = true;
            }
            Properties.Settings.Default.Save();
        }

        private void MakeScreenShotsDocument_Activated(object sender, System.EventArgs e) {

            if (Properties.Settings.Default.Maximised) {
                WindowState = FormWindowState.Maximized;
                Location = Properties.Settings.Default.Location;
                Size = Properties.Settings.Default.Size;
            }
            else if (Properties.Settings.Default.Minimised) {
                WindowState = FormWindowState.Minimized;
                Location = Properties.Settings.Default.Location;
                Size = Properties.Settings.Default.Size;
            }
            else {
                Location = Properties.Settings.Default.Location;
                Size = Properties.Settings.Default.Size;
            }
        }

        private void CreateDocument_Click(object sender, EventArgs e)
        {
            CreateWordDoc cd = new CreateWordDoc();
            cd.leftBrowser = LeftBrowser.Text.ToString();
            cd.leftScreenShotPath = LeftScreenShotsPath.Text.ToString();
            // leave cp.rightBrowser blank to have only the left side populated
            cd.rightBrowser = RightBrowser.Text.ToString();
            cd.rightScreenShotPath = RightScreenShotsPath.Text.ToString();
            cd.ExcelPathAndFile = ExcelFile.Text.ToString();
            cd.HeaderPage = HeaderPage.Text.ToString();
            cd.overrideEtension = OverrideExtension.Text.ToString();
            cd.CreateDoc();
            var result = MessageBox.Show("Document Created");
        }

        private void Help_Click(object sender, EventArgs e)
        {
            string path;
            //path = GetWebPageTitle(url);
            path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase);
            Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
            Document document = application.Documents.Open(string.Format("{0}\\Resources\\Help.docx",path), ReadOnly:true);
            application.Visible = true;
        }

        private void CreateExcelDocument_Click(object sender, System.EventArgs e)
        {
            string url = Url.Text.ToString();       // "http://www.CialisMD.com/Pages/";
            PopulateExcel(url);
            var result = MessageBox.Show("Excel File Created");
        }
    }
}
