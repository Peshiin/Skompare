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
using WinForms = System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace SkompareWPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private MainHandler MainHandler;
        private Color HighlightColor;
        private SolidColorBrush HighlightBrush = new SolidColorBrush();
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

        private void SelectColorButton_Click(object sender, RoutedEventArgs e)
        {
            WinForms.ColorDialog colorDialog = new WinForms.ColorDialog();
            colorDialog.AllowFullOpen = false;
            if(colorDialog.ShowDialog() == WinForms.DialogResult.OK)
            {
                System.Drawing.Color highlightColor = colorDialog.Color;

                int red = highlightColor.R;
                int green = highlightColor.G;
                int blue = highlightColor.B;

                bool IsLowContrast = false;

                if((red < 200 && green < 200 && blue < 200)
                    || (red < 150 && green < 150))
                {
                    IsLowContrast = true;
                }

                if(IsLowContrast)
                    SelectColorTextBox.Foreground = Brushes.White;
                else
                    SelectColorTextBox.Foreground = Brushes.Black;

                HighlightColor = Color.FromArgb(255, (byte)red, (byte)green, (byte)blue);
                HighlightBrush.Color = HighlightColor;
                SelectColorTextBox.Background = HighlightBrush;
                SelectColorTextBox.Text = red + "," + green + "," + blue;
            }
        }

        private void SelectColorTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string[] colors = (SelectColorTextBox.Text).Split(',');

            try
            {
                int red = int.Parse(colors[0]);
                int green = int.Parse(colors[1]);
                int blue = int.Parse(colors[2]);

                bool IsLowContrast = false;

                if ((red < 200 && green < 200 && blue < 200)
                    || (red < 150 && green < 150))
                {
                    IsLowContrast = true;
                }

                if (IsLowContrast)
                    SelectColorTextBox.Foreground = Brushes.White;
                else
                    SelectColorTextBox.Foreground = Brushes.Black;

                HighlightColor = Color.FromArgb(255, (byte)red, (byte)green, (byte)blue);
                HighlightBrush.Color = HighlightColor;
                SelectColorTextBox.Background = HighlightBrush;
            }
            catch (Exception ex) { }
        }

        private void StartRowTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9]+"); 
            e.Handled = regex.IsMatch(e.Text);
        }

        private void SearchColumnTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^A-Z]");
            e.Handled = regex.IsMatch(e.Text);
        }
    }
}
