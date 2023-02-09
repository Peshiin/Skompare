using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace SkompareWPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private MainHandler MainHandler;

        public MainWindow()
        {
            InitializeComponent();
            MainHandler = new MainHandler(OldFileControl, NewFileControl);
            DataContext = MainHandler;

            OldFileControl.XlFile = MainHandler.OldFile;
            NewFileControl.XlFile = MainHandler.NewFile;
        }

        private void LanguageSwitcherButton_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(OldFileControl.XlFileName + " " + NewFileControl.XlFileName);
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Excel.Application xlApp = MainHandler.XlApp;
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }
    }
}
