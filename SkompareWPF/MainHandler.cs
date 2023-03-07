﻿using SkompareWPF.Components;
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
using System.IO;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Collections;
using System.Reflection;
using System.Text.RegularExpressions;

namespace SkompareWPF
{
    internal class MainHandler
    {
        public Excel.Application XlApp{ get; private set; } = new Excel.Application();
        public XlFile OldFile { get; private set; }
        public XlFile NewFile { get; private set; }
        private Workbook ResultWorkbook { get; set; }
        private Worksheet ResultWorksheet { get; set; }
        private List<List<string>> NewList { get; set; }
        private List<List<string>> OldList { get; set; }
        private List<string> NewSearchList { get; set; }
        private List<string> OldSearchList { get; set; } 
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
                OldFile.Workbook = XlApp.Workbooks.Open(OldFile.FilePath);
                NewFile.Workbook = XlApp.Workbooks.Open(NewFile.FilePath);

                CheckColumns();

                //Assigns sheets arrays
                NewList = GetSheet2DList(NewFile.SelectedSheet, NewFile.RowsCount, NewFile.ColumnsCount);
                OldList = GetSheet2DList(OldFile.SelectedSheet, OldFile.RowsCount, OldFile.ColumnsCount);

                //Assigns search arrays
                NewSearchList = GetSearchList(NewList);
                OldSearchList = GetSearchList(OldList);

                //Sets auto updating of the XlApp to false
                //autoUpdate(False);

                //Creates "result" workbook to where the actual comparing will be done
                CreateResult();
                // Removes absolute file reference from header
                if(StartRow > 1)
                {
                    Range header = ResultWorksheet.Range["A1:" + GetExcelColumnName(NewFile.ColumnsCount) + (StartRow - 1).ToString()];
                    string headerValue;
                    string fileReference;
                    foreach (Range cell in header)
                    {
                        //C:\Users\pechm\Desktop\skompare test\[newTest.xlsx]
                        headerValue = Convert.ToString(cell.Formula);
                        fileReference = "[" + NewFile.Workbook.Name + "]"; //NewFile.Workbook.Path + "\\
                        if (headerValue == null)
                            continue;
                        Trace.WriteLine(fileReference + " " + headerValue);
                        while(headerValue.Contains(fileReference))
                        {
                            headerValue = headerValue.Replace(fileReference, "");
                            cell.Formula = headerValue;
                        }
                    }
                }

                CompareRowsCount = Math.Max(OldFile.RowsCount, NewFile.RowsCount);

                //Comparison itself
                Compare();

                //Allows auto updating
                //autoUpdate(True);

                //Closes the originals and shows the result
                if (OldFile.Workbook != NewFile.Workbook)
                {
                    OldFile.Workbook.Close(SaveChanges: false);
                    NewFile.Workbook.Close(SaveChanges: false);
                }
                else
                    OldFile.Workbook.Close(SaveChanges: false);

                XlApp.Visible = true;
            }
            catch(Exception ex)
            {
                Trace.WriteLine(ex.StackTrace +
                                Environment.NewLine +
                                ex.Message);
                Trace.Flush();
                ResultWorkbook.Close(SaveChanges: false);
                Marshal.ReleaseComObject(ResultWorkbook);
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
            string tempFilePath = Path.GetTempPath() + "SkompareTempFile" + Path.GetExtension(OldFile.FilePath);
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
            ResultWorksheet = ResultWorkbook.Worksheets[NewFile.SelectedSheet.Name];

            ResultWorkbook.Unprotect();
            ResultWorksheet.Unprotect();
        }
        
        /// <summary>
        /// Goes through the worksheets and compares rows
        /// </summary>
        private void Compare()
        {
            bool duplicityFound = false;
            string searchString;
            int oldRowIndex;

            List<bool> duplicity = new List<bool>();
            for (int i = 0; i < CompareRowsCount; i++)
                duplicity.Add(false);

            Trace.WriteLine("Starting looping");
            for(int newRowIndex = StartRow - 1; newRowIndex < NewFile.RowsCount; newRowIndex++)
            {
                Trace.WriteLine("Row index: " + newRowIndex);

                searchString = NewSearchList[newRowIndex];
                Trace.WriteLine("Searching for: " + searchString);

                oldRowIndex = OldSearchList.IndexOf(searchString);
                Trace.WriteLine("Found at row " + (oldRowIndex + 1) + " of old sheet");

                if (oldRowIndex < 0)
                {
                    ResultWorksheet.Rows[newRowIndex + 1].EntireRow.Interior.Color =
                        System.Drawing.Color.FromArgb(HighlightColor.R, HighlightColor.G, HighlightColor.B);
                    continue;
                }

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
                    else
                        oldRowIndex = OldSearchList.IndexOf(searchString, oldRowIndex + 1);
                }

                if(oldRowIndex >= 0)
                {
                    duplicity[oldRowIndex] = true;

                    CompareRow(newRowIndex, oldRowIndex);
                }
            }

            Trace.WriteLine("Deleting rows from \"Cancelled\"");
            DeleteRows(ResultWorkbook.Worksheets["Cancelled"], duplicity);
        }

        
        /// <summary>
        /// Deletes "found" rows from "Cancelled" sheet in "result" workbook
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="indexArray"></param>
        private void DeleteRows(Worksheet sheet, List<bool> indexArray)
        {
            for(int i = indexArray.Count - 1; i >= StartRow - 1; i--)
            {
                if (indexArray[i])
                    sheet.Rows[i + 1].EntireRow.Delete();
            }
        }

        
        /// <summary>
        /// Gets array of indexes to search by
        /// </summary>
        /// <param name="inputArray"></param>
        /// <returns></returns>
        private List<string> GetSearchList(List<List<string>> inputArray)
        {
            int inputLength = inputArray.Count();
            List<string> returnList = new List<string>();
            while (returnList.Count < StartRow - 1)
                returnList.Add(null);

            for(int row = StartRow - 1; row < inputLength; row++)
            {
                returnList.Add(string.Empty);

                foreach(string key in SearchColumns)
                {
                    if (key != string.Empty && key != "")
                    {
                        returnList[row] += inputArray[row][GetExcelColumnNumber(key) - 1];
                    }
                }
            }

            return returnList;
        }

        /// <summary>
        /// Returns range of Excel worksheet as List of Lists<string>
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rows"></param>
        /// <param name="columns"></param>
        /// <returns></returns>
        private List<List<string>> GetSheet2DList(Worksheet sheet, int rows, int columns)
        {
            List<List<string>> returnList = new List<List<string>>();
            List<string> rowList;

            for(int row = 1; row <= rows; row++)
            {
                rowList = new List<string>();
                for (int col = 1; col <= columns; col++)
                    if(sheet.Cells[row, col].Value != null)
                        rowList.Add(sheet.Cells[row, col].Value.ToString());
                    else
                        rowList.Add(null);
                returnList.Add(rowList);
            }

            return returnList;
        }

        
        /// <summary>
        /// Compares values in single rows
        /// </summary>
        /// <param name="NewR"></param>
        /// <param name="OldR"></param>
        private void CompareRow(int newR, int oldR)
        {
            string newVal = null;
            string oldVal = null;
            Range row = ResultWorksheet.Rows[newR + 1];

            try
            {
                for(int column = 0; column < NewFile.ColumnsCount; column++)
                {
                    if (NewList[newR][column] != null)
                        newVal = NewList[newR][column].ToString();

                    if (OldList[oldR][column] != null)
                        oldVal = OldList[oldR][column].ToString();

                    if (newVal == null && oldVal == null)
                        continue;
                    else if(newVal == null && oldVal != null)
                        CompareStyle(row.Cells[column + 1], newVal, oldVal);
                    else if(!newVal.Equals(oldVal, StringComparison.InvariantCultureIgnoreCase))
                        CompareStyle(row.Cells[column + 1], newVal, oldVal);
                }
            }
            catch(Exception ex)
            {
                Trace.WriteLine(ex.StackTrace);
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
            if(newStr == null)
                newRng.NumberFormat = "General";
            else if (newStr.Contains("."))
                // sets range format to "Text"
                newRng.NumberFormat = "@";
            else
                newRng.NumberFormat = "General";

            if(ChangesHighlight == "HighlightOnlyRadioButton")
            {
                newRng.Interior.Color
                    = System.Drawing.Color.FromArgb(HighlightColor.R, HighlightColor.G, HighlightColor.B);
                newRng.Value = newStr;
            }

            else if(ChangesHighlight == "HighlightCommentRadioButton")
            {
                newRng.Interior.Color
                    = System.Drawing.Color.FromArgb(HighlightColor.R, HighlightColor.G, HighlightColor.B);
                if(newStr != null)
                    newRng.Value = newStr;

                if (newRng.Comment != null)
                    newRng.Comment.Delete();

                if (oldStr == null || oldStr == "" || oldStr == string.Empty)
                    newRng.AddComment("-");
                else
                    newRng.AddComment(StartString + oldStr + EndString);
            }

            else if(ChangesHighlight == "HighlightStringRadioButton")
            {
                newRng.Interior.Color
                    = System.Drawing.Color.FromArgb(HighlightColor.R, HighlightColor.G, HighlightColor.B);
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

        /// <summary>
        /// Gets Excel column number from letter 
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private int GetExcelColumnNumber(string name)
        {
            int number = 0;
            int pow = 1;
            for (int i = name.Length - 1; i >= 0; i--)
            {
                number += (name[i] - 'A' + 1) * pow;
                pow *= 26;
            }

            return number;
        }

        private string GetExcelColumnName(int columnNumber)
        {
            string columnName = "";

            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }

            return columnName;
        }
    }
}
