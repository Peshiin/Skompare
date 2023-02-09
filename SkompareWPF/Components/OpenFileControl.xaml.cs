using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace SkompareWPF.Components
{
    /// <summary>
    /// Interaction logic for OpenFileControl.xaml
    /// </summary>
    public partial class OpenFileControl : UserControl, INotifyPropertyChanged, INotifyCollectionChanged
    {
        public string Label
        {
            get { return (string)GetValue(LabelProperty); }
            set { SetValue(LabelProperty, value); }
        }

        // Using a DependencyProperty as the backing store for Label.
        public static readonly DependencyProperty LabelProperty =
            DependencyProperty.Register("Label", typeof(string), typeof(OpenFileControl), new PropertyMetadata(string.Empty));


        // Using a DependencyProperty as the backing store for ControlHeight.
        public static readonly DependencyProperty ControlHeightProperty =
            DependencyProperty.Register("ControlHeight", typeof(double), typeof(OpenFileControl), new PropertyMetadata(0.0));

        public XlFile XlFile
        {
            get { return (XlFile)GetValue(XlFileProperty); }
            set { SetValue(XlFileProperty, value); }
        }

        // Using a DependencyProperty as the backing store for XlFile.
        public static readonly DependencyProperty XlFileProperty =
            DependencyProperty.Register("XlFile", typeof(XlFile), typeof(OpenFileControl), new PropertyMetadata(null));


        public event PropertyChangedEventHandler PropertyChanged;
        public event NotifyCollectionChangedEventHandler CollectionChanged;

        protected void InvokeChange(string property)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(property));
        }

        public string xlFileName;
        /// <summary>
        /// Selected excel file name
        /// </summary>
        public string XlFileName
        {
            get
            {
                return xlFileName;
            }
            set
            {
                xlFileName = value;
                FileTextBox.Text = value;                
                InvokeChange(XlFileName);
                XlFile.Worksheets.CollectionChanged += Worksheets_CollectionChanged;
            }
        }

        public OpenFileControl()
        {
            InitializeComponent();
            DataContext = this;
            SheetComboBox.SelectedIndex = 0;
        }

        private void Worksheets_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            SheetComboBox.SelectedIndex = 0;
        }

        private void FileOpenerButton_Click(object sender, RoutedEventArgs e)
        {            
            var dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.Title = "Select file";
            dialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            if (dialog.ShowDialog() == true)
            {
                XlFileName = dialog.FileName;
                FileTextBox.Text = xlFileName;
            }
        }
        
        private void FileTextBox_PreviewDrop(object sender, DragEventArgs e)
        {
            var fileNames = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (fileNames == null) return;

            var fileName = fileNames.FirstOrDefault();
            if (fileName == null) return;

            if (fileName.Contains("xls"))
                XlFileName = fileName;
            else
                throw new Exception("Invalid file type entered");
        }

        private void FileTextBox_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Handled = true;
        }
    }
}
