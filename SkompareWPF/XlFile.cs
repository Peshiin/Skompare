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
            private set
            {
                workbook = value;
                Worksheets.Clear();
                foreach (Excel.Worksheet sheet in Workbook.Worksheets)
                    Worksheets.Add(sheet);
                InvokeChange(nameof(Worksheets));
            }
        }
        public ObservableCollection <Worksheet> Worksheets { get ; private set; }
        public Excel.Worksheet SelectedSheet { get; set; }
        private string FilePath { get; set; }
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
                Workbook = XlApp.Workbooks.Open(FilePath);
                InvokeChange(nameof(Workbook));
            }
                
        }
    }
}
