using SkompareWPF.Components;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Media;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using static System.Net.Mime.MediaTypeNames;
using System.IO;
using static System.Net.WebRequestMethods;
using System.Windows.Input;
using System.Globalization;

namespace SkompareWPF
{
    internal class MainHandler
    {
        public Excel.Application XlApp{ get; private set; } = new Excel.Application();
        public XlFile OldFile { get; private set; }
        public XlFile NewFile { get; private set; }
        private Workbook ResultWorkbook { get; set; }
        private Worksheet ResultWorksheet { get; set; }
        private object[,] NewArray { get; set; }
        private object[,] OldArray { get; set; }
        private string[] NewSearchArr { get; set; }
        private string[] OldSearchArr { get; set; } 
        public Color HighlightColor { get; set; }
        public string ChangesHighlight { get; set; } = string.Empty;
        public List<string> SearchColumns { get; set; } = new List<string>(3);
        private int CompareRowsCount { get; set; }
        public int StartRow { get; set; }
        public string StartString { get; set; }
        public string EndString { get; set; } 

        public MainHandler(OpenFileControl oldControl, OpenFileControl newControl)
        {
            OldFile = new XlFile(oldControl, XlApp);
            NewFile = new XlFile(newControl, XlApp);

            for(int i = 1; i <= 3; i++)
                SearchColumns.Add(string.Empty);
        }

        public void CompareInit()
        {
            string debugFilePath = AppDomain.CurrentDomain.BaseDirectory + "\\Debug.log";
            if(System.IO.File.Exists(debugFilePath))
            {
                try
                {
                    System.IO.File.Delete(debugFilePath);
                }
                catch(Exception ex)
                {
                    Trace.WriteLine(ex.Message);
                }
            }

            TextWriterTraceListener debug = new TextWriterTraceListener(debugFilePath, "myListener");
            Trace.Listeners.Add(debug);
            Trace.WriteLine("Starting comparing @ " + DateTime.Now.ToString());
            Trace.Indent();

            try
            {
                CheckColumns();

                //Assigns sheets arrays
                NewArray = GetSheetArray(NewFile.SelectedSheet, NewFile.RowsCount, NewFile.ColumnsCount);
                OldArray = GetSheetArray(OldFile.SelectedSheet, OldFile.RowsCount, OldFile.ColumnsCount);

                //Assigns search arrays
                NewSearchArr = GetSearchArray(NewArray);
                OldSearchArr = GetSearchArray(OldArray);

                //Sets auto updating of the XlApp to false
                //autoUpdate(False);

                //Creates "result" workbook to where the actual comparing will be done
                CreateResult();

                CompareRowsCount = Math.Max(OldFile.RowsCount, NewFile.RowsCount);

                //Comparison itself
                Compare();

                //Allows auto updating
                //autoUpdate(True);

                //Closes the originals and shows the result
                OldFile.Workbook.Close(SaveChanges: false);
                NewFile.Workbook.Close(SaveChanges: false);
                XlApp.Visible = true;
            }
            catch(Exception ex)
            {
                Trace.WriteLine(ex.StackTrace +
                                Environment.NewLine +
                                ex.Message);
                Trace.Flush();
            }

            Trace.Unindent();
            Trace.WriteLine("Comparing ended");
            Trace.WriteLine("___________________________________________________");
            Trace.Flush();
        }

        
        /// <summary>
        /// Checks the number of columns in both sheets and lets the user to unite this
        /// </summary>
        private void CheckColumns()
        {
            if(NewFile.ColumnsCount != OldFile.ColumnsCount)
            {
                if (MessageBox.Show("Rozdílný počet sloupců ve vybraných listech." +
                                    Environment.NewLine +
                                    "Chcete se pokusit o úpravu?",
                                    "Close",
                                    MessageBoxButtons.YesNo,
                                    MessageBoxIcon.Question) 
                    == System.Windows.Forms.DialogResult.Yes)
                {
                    MessageBox.Show("Následně se otevře aplikace pro přidání sloupců" +
                                    Environment.NewLine +
                                    "Přidejte sloupce tam, kde chybí, ale" +
                                    Environment.NewLine +
                                    Environment.NewLine +
                                    "!!! NEZAVÍREJTE OKNO EXCELU !!!");

                    XlApp.Visible = true;

                    MessageBox.Show("Hotovo?" +
                                    Environment.NewLine + 
                                    "Doplnili jste všechny sloupce na správná místa?",
                                    "Hotovo?",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Question,
                                    MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.DefaultDesktopOnly);

                    XlApp.Visible = false;

                    OldFile.SelectedSheet = OldFile.Workbook.Worksheets[OldFile.SelectedSheet.Index];
                    NewFile.SelectedSheet = NewFile.Workbook.Worksheets[NewFile.SelectedSheet.Index];
                }
            }
        }
        
        /// <summary>
        /// Creates "result" workbook
        /// </summary>
        /// <returns></returns>
        private void CreateResult()
        {
            string tempFilePath = Path.GetTempPath() + "\\SkompareTempFile" + Path.GetExtension(OldFile.FilePath);
            OldFile.Workbook.SaveCopyAs(tempFilePath);
            Trace.WriteLine("Created result as temporary: " + tempFilePath);

            ResultWorkbook = XlApp.Workbooks.Open(tempFilePath);

            try
            {
                ResultWorkbook.Worksheets["Cancelled"].Delete();
            }
            catch(Exception ex)
            {
                Trace.WriteLine("\"Cancelled\" sheet does not exist");
            }
            ResultWorkbook.Worksheets[OldFile.SelectedSheet.Index].Name = "Cancelled";
            NewFile.SelectedSheet.Copy(Before: ResultWorkbook.Worksheets["Cancelled"]);

            ResultWorkbook.Unprotect();
            ResultWorksheet.Unprotect();
        }
        
        /// <summary>
        /// Goes through the worksheets and compares rows
        /// </summary>
        private void Compare()
        {
            bool matchFound = false;
            bool duplicityFound = false;
            string searchString = string.Empty;
            int oldRowIndex = -1;

            bool[] duplicity = new bool[CompareRowsCount];
            for (int i = 0; i <= duplicity.Count(); i++)
                duplicity[i] = false;

            Trace.WriteLine("Starting looping");

            for(int newRowIndex = StartRow; newRowIndex < NewFile.RowsCount; newRowIndex++)
            {
                Trace.WriteLine("Row index: " + newRowIndex);
                matchFound = false;
                searchString = NewSearchArr[newRowIndex];

                oldRowIndex = Array.IndexOf(OldSearchArr, searchString);

                if (oldRowIndex <= 0)
                    continue;

                if (duplicity[oldRowIndex])
                {
                    if (!duplicityFound)
                    {
                        MessageBox.Show("Nalezena duplicita zadaných vyhledávacích klíčů" +
                                        Environment.NewLine +
                                        "Skript proběhne s předpokladem max. dvou duplicit." +
                                        Environment.NewLine +
                                        "Pokud je předpokládané množství duplicit více, ošetřete vhodným výběrem klíčů.");
                        duplicityFound = true;
                    }

                    oldRowIndex = Array.IndexOf(OldSearchArr, searchString, oldRowIndex + 1);
                }

                if(oldRowIndex > 0)
                {
                    matchFound = true;
                    duplicity[oldRowIndex] = true;

                    CompareRow(newRowIndex, oldRowIndex);
                }

                if (!matchFound)
                    ResultWorksheet.Rows.EntireRow.Interior.Color = HighlightColor;
            }

            Trace.WriteLine("Deleting rows from \"Cancelled\"");
            DeleteRows(ResultWorkbook.Worksheets["Cancelled"], duplicity);
        }

        
        /// <summary>
        /// Deletes "found" rows from "Cancelled" sheet in "result" workbook
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="indexArray"></param>
        private void DeleteRows(Worksheet sheet, bool[] indexArray)
        {
            for(int i = indexArray.Length - 1; i >= StartRow; i--)
            {
                if (indexArray[i])
                    sheet.Rows[i].EntireRow.Delete();
            }
        }

        
        /// <summary>
        /// Gets array of indexes to search by
        /// </summary>
        /// <param name="inputArray"></param>
        /// <returns></returns>
        private string[] GetSearchArray(Object[,] inputArray)
        {
            int inputLength = inputArray.GetLength(1);
            string[] returnArray = new string[inputLength];

            for(int row = StartRow; row <= inputLength; row++)
            {
                foreach(string key in SearchColumns)
                {
                    if (key != string.Empty && key != "")
                        returnArray[row] += key;
                }
            }

            return returnArray;
        }        
        
        /// <summary>
        /// Gets array of values from range to be compared
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rows"></param>
        /// <param name="columns"></param>
        /// <returns></returns>
        private object[,] GetSheetArray(Worksheet sheet, int rows, int columns)
        {
            string lastCell = columns.ToString() + rows.ToString();

            return (Object[,])sheet.Range["A1", lastCell].Value;
        }

        
        /// <summary>
        /// Compares values in single rows
        /// </summary>
        /// <param name="NewR"></param>
        /// <param name="OldR"></param>
        private void CompareRow(int newR, int oldR)
        {
            string newVal;
            string oldVal;
            Range row = ResultWorksheet.Rows[newR];

            try
            {
                for(int column = 1; column <= NewFile.ColumnsCount; column++)
                {
                    newVal = NewArray.GetValue(newR, column) as string;
                    oldVal = OldArray.GetValue(oldR, column) as string;

                    if (string.Compare(newVal, oldVal, true, CultureInfo.InvariantCulture) == 0)
                        CompareStyle(row.Cells[1, column], newVal, oldVal);
                }
            }
            catch(Exception ex)
            {
                Trace.WriteLine(ex.Message);
            }
        }

        
        /// <summary>
        /// Defines how the differences found shall be highlighted
        /// </summary>
        /// <param name="newRng"></param>
        /// <param name="newStr"></param>
        /// <param name="oldStr"></param>
        private void CompareStyle(Range newRng, string newStr, string oldStr)
        {
            // sets range format to "Text"
            if(newStr.Contains("."))
                newRng.NumberFormat = "@";
            else
                newRng.NumberFormat = "General";

            if(ChangesHighlight == "HighlightOnlyRadioButton")
            {
                newRng.Interior.Color = HighlightColor;
                newRng.Value = newStr;
            }

            else if(ChangesHighlight == "HighlightCommentRadioButton")
            {
                newRng.Interior.Color = HighlightColor;
                newRng.Value = newStr;

                if (newRng.Comment != null)
                    newRng.Comment.Delete();

                if (oldStr == "" || oldStr == string.Empty)
                    newRng.AddComment("-");
                else
                    newRng.AddComment(StartString + oldStr + EndString);
            }

            else if(ChangesHighlight == "HighlightStringRadioButton")
            {
                newRng.Interior.Color = HighlightColor;
                newRng.Value = newStr + " " + StartString + oldStr + EndString;
            }

            else if(ChangesHighlight == "CommentOnlyRadioButton")
            {
                newRng.Value = newStr;

                if (newRng.Comment != null)
                    newRng.Comment.Delete();

                if (oldStr == "" || oldStr == string.Empty)
                    newRng.AddComment("-");
                else
                    newRng.AddComment(StartString + oldStr + EndString);
            }

            else if (ChangesHighlight == "StringOnlyRadioButton")
            {
                newRng.Value = newStr + " " + StartString + oldStr + EndString;
            }
        }
    }
}
