using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
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
using WinForms = System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using SkompareWPF.Components;

namespace SkompareWPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private MainHandler MainHandler;
        private SolidColorBrush HighlightBrush = new SolidColorBrush();
        private string selectedRadioButton = string.Empty;

        public MainWindow()
        {
            InitializeComponent();
            MainHandler = new MainHandler(OldFileControl, NewFileControl);
            DataContext = MainHandler;

            OldFileControl.XlFile = MainHandler.OldFile;
            NewFileControl.XlFile = MainHandler.NewFile;

            MainHandler.HighlightColor = ((SolidColorBrush)SelectColorTextBox.Background).Color;
            MainHandler.ChangesHighlight = selectedRadioButton;
            MainHandler.StartRow = int.Parse(StartRowTextBox.Text);
            MainHandler.StartString = StartStringTextBox.Text;
            MainHandler.EndString = EndStringTextBox.Text;
            MainHandler.SearchColumns[0] = SearchColumnATextBox.Text;
            MainHandler.PropertyChanged += MainHandler_PropertyChanged;
        }

        private void MainHandler_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
        }

        private void LanguageSwitcherButton_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(OldFileControl.XlFileName + " " + NewFileControl.XlFileName);
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                if(MainHandler.OldFile != null)
                {
                    MainHandler.OldFile.Workbook.Close(SaveChanges: false);
                }
                if(MainHandler.NewFile != null)
                {
                    MainHandler.NewFile.Workbook.Close(SaveChanges: false);
                }
            }
            catch(Exception ex)
            {
                Trace.WriteLine(ex.Message);
            }
            MainHandler.XlApp.Quit();
            Marshal.ReleaseComObject(MainHandler.XlApp);
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

                MainHandler.HighlightColor = Color.FromArgb(255, (byte)red, (byte)green, (byte)blue);
                HighlightBrush.Color = MainHandler.HighlightColor;
                SelectColorTextBox.Background = HighlightBrush;
                SelectColorTextBox.Text = red + "," + green + "," + blue;
            }
        }

        private void SelectColorTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string[] colors = (SelectColorTextBox.Text).Split(',');
            
            if(MainHandler != null)
            {
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

                    MainHandler.HighlightColor = Color.FromArgb(255, (byte)red, (byte)green, (byte)blue);
                    HighlightBrush.Color = MainHandler.HighlightColor;
                    SelectColorTextBox.Background = HighlightBrush;
                }
                catch (Exception ex)
                {
                    Trace.WriteLine(ex.ToString());
                }
            }
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

        private void ChangesHighlightRadioButton_Checked(object sender, RoutedEventArgs e)
        {
            if (MainHandler != null)
                MainHandler.ChangesHighlight = (e.Source as RadioButton).Name;
            else
                selectedRadioButton = (e.Source as RadioButton).Name;
        }

        private void StartRowTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                if (MainHandler == null)
                    return;
                if ((e.Source as TextBox).Text == null)
                    return;

                MainHandler.StartRow = int.Parse((e.Source as TextBox).Text);
            }
            catch(Exception) { }
        }

        private void StringTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (MainHandler == null)
                return;
            else if ((e.Source as TextBox).Name == nameof(StartStringTextBox))
                MainHandler.StartString = (e.Source as TextBox).Text;
            else if ((e.Source as TextBox).Name == nameof(EndStringTextBox))
                MainHandler.EndString = (e.Source as TextBox).Text;
        }

        private void SearchColumnTextBox_Changed(object sender, TextChangedEventArgs e)
        {
            if(MainHandler == null)
                return ;

            else if ((e.Source as TextBox).Name == nameof(SearchColumnATextBox))
                MainHandler.SearchColumns[0] = (e.Source as TextBox).Text;
            else if ((e.Source as TextBox).Name == nameof(SearchColumnBTextBox))
                MainHandler.SearchColumns[1] = (e.Source as TextBox).Text;
            else if ((e.Source as TextBox).Name == nameof(SearchColumnCTextBox))
                MainHandler.SearchColumns[2] = (e.Source as TextBox).Text;
        }

        private void SearchColumnCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            if (MainHandler == null)
                return;

            else if ((e.Source as CheckBox).Name == nameof(SearchColumnBCheckBox))
            {
                if ((e.Source as CheckBox).IsChecked == true)
                    MainHandler.SearchColumns[1] = SearchColumnBTextBox.Text;
                else
                    MainHandler.SearchColumns[1] = string.Empty;
            }
            else if ((e.Source as CheckBox).Name == nameof(SearchColumnCCheckBox))
            {
                if ((e.Source as CheckBox).IsChecked == true)
                    MainHandler.SearchColumns[2] = SearchColumnCTextBox.Text;
                else
                    MainHandler.SearchColumns[2] = string.Empty;
            }
        }

        private void StartCompareButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (MainHandler.NewFile.Workbook == null || MainHandler.OldFile.Workbook == null)
                    throw new Exception("Nebyl správně vybrán porovnávaný sešit");

                if (Path.GetExtension(MainHandler.OldFile.FilePath) !=
                    Path.GetExtension(MainHandler.NewFile.FilePath))
                    throw new Exception("Formáty souborů se neshodují. Porovnávejte soubory se stejným formátem");

                MainHandler.CompareInit();
            }
            catch(Exception ex)
            {
                Trace.WriteLine(ex.ToString());     
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
