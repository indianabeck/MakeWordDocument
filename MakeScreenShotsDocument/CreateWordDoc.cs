using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace MakeScreenShotsDocument
{
    public class CreateWordDoc
    {
        #region private properties
        private string _leftBrowser = string.Empty;
        private string _leftScreenShotPath = string.Empty;
        private string _rightBrowser = string.Empty;
        private string _rightScreenShotPath = string.Empty;
        private string _ExcelPathAndFile = string.Empty;
        private string _HeaderPage = string.Empty;
        private string _overrideExtension = string.Empty;
        private class _screenImage
        {
            public string LeftPath { get; set; }
            public string RightPath { get; set; }
            public string Title { get; set; }
            public string Name { get; set; }
        }
        #endregion

        #region public propeties
        public string leftBrowser
        {
            get { return _leftBrowser; }
            set { _leftBrowser = value; }
        }
        public string leftScreenShotPath
        {
            get { return _leftScreenShotPath; }
            set { _leftScreenShotPath = value; }
        }
        public string rightBrowser
        {
            get { return _rightBrowser; }
            set { _rightBrowser = value; }
        }
        public string rightScreenShotPath
        {
            get { return _rightScreenShotPath; }
            set { _rightScreenShotPath = value; }
        }

        public string ExcelPathAndFile
        {
            get { return _ExcelPathAndFile; }
            set { _ExcelPathAndFile = value; }
        }

        public string HeaderPage
        {
            get { return _HeaderPage; }
            set { _HeaderPage = value; }
        }

        public string overrideEtension
        {
            get { return _overrideExtension; }
            set { _overrideExtension = value; }
        }
        #endregion

        object oMissing = System.Reflection.Missing.Value;
        object oTrueValue = true;
        Object oFalseValue = false;
        // predefined bookmarks 
        object oEndOfDoc = "\\endofdoc";
        object oStartOfDoc = "\\StartOfDoc";
        object styleHeading1 = "Heading 1";

        public void CreateDoc()
        {
            string leftFilePath = leftScreenShotPath;
            string rightFilePath = rightScreenShotPath;

            #region fill screenImages List from Excel file
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook;
            try {
                xlWorkbook = xlApp.Workbooks.Open(ExcelPathAndFile);
            }
            catch (Exception ex) {
                Console.WriteLine(String.Format("Could not open file {0}", ExcelPathAndFile));
                throw;
            }

            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            List<_screenImage> screenImages = new List<_screenImage>();
            if (!leftFilePath.EndsWith(@"\")) {
                leftFilePath += @"\";
            }
            if (!rightFilePath.EndsWith(@"\") && rightFilePath != "") {
                rightFilePath += @"\";
            }

            for (var i = 2; i <= rowCount; i++) {
                //test for emppty Overide (column 1, cell 1) for sheets where the "UsedRange" includes unused cells
                string tester = string.Empty;
                try {
                    tester = xlRange.Cells[i, 1].Value2.ToString();
                }
                catch (Exception ex) {
                    tester = "";
                }
                if (tester == "") {
                    break;
                }
                _screenImage singleScreen = new _screenImage();
                for (int j = 1; j <= colCount; j++) {
                    switch (j) {
                        case 2:
                            singleScreen.LeftPath = string.Format("{0}{1}", leftFilePath, xlRange.Cells[i, j].Value2.ToString());
                            singleScreen.RightPath = string.Format("{0}{1}", rightFilePath, xlRange.Cells[i, j].Value2.ToString());
                            string[] words = xlRange.Cells[i, j].Value2.ToString().Split('_');
                            singleScreen.Name = String.Format("{0}.aspx", words[0]);
                            screenImages.Add(singleScreen);
                            break;
                        case 3:
                            singleScreen.Title = xlRange.Cells[i, j].Value2.ToString();
                            break;
                        default:
                            break;
                    }
                }
            }
            xlWorkbook.Close();
            xlApp.Quit();
            #endregion

            //Start Word and create a new document.
            Word._Application oWord;
            Word._Document oDoc;

            oWord = new Word.Application();
            oWord.Visible = true;

            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
                ref oMissing, ref oMissing);
            oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
            // 1" = 72 points
            oDoc.PageSetup.LeftMargin = (int)(.5 * 72);
            oDoc.PageSetup.RightMargin = (int)(.5 * 72);
            oDoc.PageSetup.TopMargin = (int)(.5 * 72);
            oDoc.PageSetup.BottomMargin = (int)(.5 * 72);

            oDoc.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter;
            oDoc.ActiveWindow.Selection.TypeText(string.Format("{0} Screen Images \t Page ", leftBrowser));
            Object TotalPages = Word.WdFieldType.wdFieldNumPages;
            Object CurrentPage = Word.WdFieldType.wdFieldPage;
            oDoc.ActiveWindow.Selection.Fields.Add(oDoc.ActiveWindow.Selection.Range, ref CurrentPage, ref oMissing, ref oMissing);
            oDoc.ActiveWindow.Selection.TypeText(" of ");
            oDoc.ActiveWindow.Selection.Fields.Add(oDoc.ActiveWindow.Selection.Range, ref TotalPages, ref oMissing, ref oMissing);

            oDoc.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument;

            //start toc
            oDoc.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
            Word.TableOfContents toc;
            toc = AddTOC(ref oDoc);
            oDoc.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
            //end toc

            bool firstPage = true;
            string lastTitle = "$$$$";
            int subPage = 0;
            foreach (_screenImage screenImage in screenImages) {
                if (firstPage != true) {
                    //new page
                    oDoc.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
                }
                firstPage = false;

                //Insert a paragraph at the beginning of the document.
                Word.Paragraph oPara1;
                oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);













    
                string pageTitle = screenImage.Title ?? "";
                string pageName = screenImage.Name.ToString().TrimEnd().TrimStart() ?? "";
                if (pageTitle != lastTitle) {
                    pageTitle = string.Format("{0}", pageTitle);
                    lastTitle = pageTitle;
                    subPage = 0;
                }
                else { pageTitle = ""; }
                subPage++;
                oPara1.Range.Text = pageTitle;
                oPara1.Range.Font.Bold = 1;

                //add heading style with paragraph
                oPara1.Range.set_Style(ref styleHeading1);

                oPara1.Format.SpaceAfter = 24;    //24 pt spacing after paragraph.l
                oPara1.Range.InsertParagraphAfter();

                //insert an image
                object oImageRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                object saveWithDocument = true;
                object missing = Type.Missing;

                Word.Table imageTable;
                Word.Range wordRange = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                imageTable = oDoc.Content.Tables.Add(wordRange, 2, 2);
                //0 = no borders
                imageTable.Borders.Enable = 0;
                imageTable.Cell(1, 1).Range.Text = string.Format("{0} ({1})", pageName, subPage.ToString());
                imageTable.Cell(1, 1).Range.Font.Size = 13;
                imageTable.Cell(1, 2).Range.Text = "";

                oImageRng = imageTable.Cell(2, 1).Range;
                string pictureName = screenImage.LeftPath;

                try {
                    var shape = oDoc.InlineShapes.AddPicture(pictureName, ref missing, ref saveWithDocument, ref oImageRng);
                    shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                    double dWidth = System.Convert.ToDouble(shape.Width);
                    // The scaling of the image doesn't really matter when we place the image
                    // into a table cell. The Image will scale automatically to the maximum
                    // size that will fit in the table cell based the table data
                    float newWidth = (float)(dWidth * 1);
                    shape.Width = newWidth;
                }
                catch (Exception ex) {
                    imageTable.Cell(2, 1).Range.Text = string.Format("{0} missing", pictureName);
                }

                if (rightBrowser != "") {
                    string pictureName1 = screenImage.RightPath.Replace(leftBrowser, rightBrowser);
                    if (overrideEtension != "") {
                        string oldExtension = Path.GetExtension(pictureName1);
                        string newExtension = overrideEtension;
                        if(!newExtension.StartsWith(".")) {
                            newExtension = "." + overrideEtension;
                        }
                        pictureName1 = pictureName1.Replace(oldExtension, newExtension);
                    }
                    object oImageRng1 = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    oImageRng1 = imageTable.Cell(2, 2).Range;
                    try {
                        var shape1 = oDoc.InlineShapes.AddPicture(pictureName1, ref missing, ref saveWithDocument, ref oImageRng1);
                        shape1.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                        double dWidth1 = System.Convert.ToDouble(shape1.Width);
                        float newWidth1 = (float)(dWidth1 * 1);
                        shape1.Width = newWidth1;
                    }
                    catch (Exception ex) {
                        imageTable.Cell(2, 2).Range.Text = string.Format("{0} missing", pictureName1);
                    }
                }
            }

            toc.Update();
            if (HeaderPage != "") {
                Word.Range startOfDocRange = oDoc.Bookmarks.get_Item(ref oStartOfDoc).Range;
                Word.Paragraph oFirstParagraph = oDoc.Content.Paragraphs[1];
                try {
                    oFirstParagraph.Range.InsertFile(HeaderPage);
                }
                catch (Exception ex) {
                    Console.WriteLine(String.Format("Could not open file {0}", HeaderPage));
                    throw;
                }
            }
        }

        private Word.TableOfContents AddTOC(ref  Word._Document oDoc)
        {
            Word.Range myRange = oDoc.Range(ref oMissing, ref oMissing);
            object oStyleName = styleHeading1;
            myRange.set_Style(ref oStyleName);
            object start = oDoc.Content.End - 1;

            Word.Range rangeForTOC = oDoc.Range(ref start, ref oMissing);
            Word.TableOfContents toc = oDoc.TablesOfContents.Add(rangeForTOC,
                ref oTrueValue, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oTrueValue,
                ref oTrueValue, ref oTrueValue, ref oTrueValue,
                ref oTrueValue, ref oTrueValue);
            return toc;
        }
    }
}
