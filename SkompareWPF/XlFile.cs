using Microsoft.Office.Interop.Excel;
using SkompareWPF.Components;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO.IsolatedStorage;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using Excel = Microsoft.Office.Interop.Excel;

namespace SkompareWPF
{
    public class XlFile : INotifyPropertyChanged
    {
        private Excel.Application XlApp;
        private Excel.Workbook workbook;
        public Excel.Workbook Workbook
        {
            get
            {
                return workbook;
            }
            set
            {
                workbook = value;
                Worksheets.Clear();
                foreach (Excel.Worksheet sheet in Workbook.Worksheets)
                    Worksheets.Add(sheet);
                InvokeChange(nameof(Worksheets));
            }
        }
        public ObservableCollection <Worksheet> Worksheets { get ; private set; }
        private Worksheet selectedSheet;
        public Excel.Worksheet SelectedSheet
        {
            get
            {
                return selectedSheet;
            }
            set
            {
                selectedSheet = value;
                try
                {
                    RowsCount = GetLast(SelectedSheet, XlSearchOrder.xlByColumns);
                    ColumnsCount = GetLast(SelectedSheet, XlSearchOrder.xlByRows);
                    Trace.WriteLine("Rows: " + RowsCount + " Columns: " + ColumnsCount);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    RowsCount = 0;
                    ColumnsCount = 0;
                }
                InvokeChange(nameof(SelectedSheet));
                InvokeChange(nameof(RowsCount));
                InvokeChange(nameof(ColumnsCount));
            }
        }
        public int RowsCount { get; private set; }
        public int ColumnsCount { get; private set; }
        public string FilePath { get; private set; }
        private OpenFileControl Control { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void InvokeChange(string property)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(property));
        }

        public XlFile(OpenFileControl control, Excel.Application xlApp)
        {
            Control = control;
            Control.PropertyChanged += Control_PropertyChanged;

            XlApp = xlApp;

            Worksheets = new ObservableCollection<Worksheet>();
        }

        private void Control_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == (sender as OpenFileControl).XlFileName) // For some reason e.PropertyName returns value
                FilePath = (sender as OpenFileControl).XlFileName;

            if(Workbook != null)
            {
                try
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    Workbook.Close(SaveChanges: false);
                    Marshal.ReleaseComObject(Workbook);
                    Marshal.ReleaseComObject(workbook);
                }
                catch(Exception ex)
                {
                    Trace.WriteLine(ex.ToString());
                }
            }

            if(FilePath != null)
            {
                Workbook = XlApp.Workbooks.Open(FilePath, ReadOnly: true);
                InvokeChange(nameof(Workbook));
            }
        }

        /// <summary>
        /// Gets number of last cell in column/row.
        /// Use xlByColumns for last row / xlByRows for last column
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="order"></param>
        /// <returns>integer number of last row/column position</returns>
        /// <exception cref="Exception"></exception>
        private int GetLast(Worksheet sheet, XlSearchOrder order)
        {
            if(sheet == null)
                return -1;

            Range last;

            last = sheet.Cells.Find(What: "*",
                                      After: sheet.Range["A1"],
                                      LookIn: Excel.XlFindLookIn.xlValues, //Excel.XlFindLookIn.xlFormulas,   
                                      LookAt: Excel.XlLookAt.xlPart,
                                      SearchOrder: order,
                                      SearchDirection: Excel.XlSearchDirection.xlPrevious,
                                      MatchCase: false);

            if( last == null )
            {
                throw new Exception ("Vybraný list (" + sheet.Name + ") je pravděpodobně prázdný.");
            }

            //Looking for last row
            if(order == Excel.XlSearchOrder.xlByColumns)
                {
                    if (last.Row < sheet.UsedRange.Rows.Count)
                        return sheet.UsedRange.Rows.Count;
                    else
                        return last.Row;
                }

            //looking for last column
            else if(order == Excel.XlSearchOrder.xlByRows)
                {
                    if (last.Column < sheet.UsedRange.Columns.Count)
                        return sheet.UsedRange.Columns.Count;
                    else
                        return last.Column;
                }

            return 0;
        }
    }
}
