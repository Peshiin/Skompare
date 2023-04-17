using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Comos.Controls;
using Comos.Global;
using Comos.Global.AppControls;
using ComosROUtilities;
using REPORTLib;
using Plt;
using ComosVBInterface;
using ComosProjTreeV;
using System.Diagnostics;

namespace LayeredTemplatesLib
{
    /// <summary>
    /// Interakční logika pro UserControl1.xaml
    /// </summary>
    public partial class MainWindow : IComosControl
    {
        private MainHandler mainHandler = new MainHandler();
        private Point mouseClickPoint;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = mainHandler;
        }

        public IComosDWorkset Workset { get; set; }
        public IComosDGeneralCollection Objects { get; set; }
        public string Parameters { get; set; }
        public IContainer ControlContainer { get; set; }

        public void OnCanExecute(CanExecuteRoutedEventArgs e)
        {
        }

        public void OnExecuted(ExecutedRoutedEventArgs e)
        {
        }

        public void OnPreviewExecuted(ExecutedRoutedEventArgs e)
        {
        }

        private void TemplateSelectionComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            IComosDProject selectedTemplateProject = null;
            IComosDCollection overlayDevices = null;

            foreach (IComosDWorkingOverlay overlay in mainHandler.TemplateProjects)
            {
                if (overlay.FullName() == TemplateSelectionComboBox.SelectedItem.ToString())
                {
                    selectedTemplateProject = overlay.Project();
                    selectedTemplateProject.let_CurrentWorkingOverlay(overlay);
                    overlayDevices = selectedTemplateProject.Devices();
                    break;
                }
            }

            if (selectedTemplateProject == null)
                throw new Exception("Chyba: Vybraná vrstva nenalezena");

            mainHandler.GetProjectTreeView(mainHandler.TemplateRootNodes, selectedTemplateProject);
        }

        private void TreeViewItem_MouseMove(object sender, MouseEventArgs e)
        {
            Point mousePosition = e.GetPosition(null);
            Vector positionDifference = mouseClickPoint - mousePosition;

            var item = Mouse.DirectlyOver;
            if (e.LeftButton == MouseButtonState.Pressed
                && item != null
                && Math.Abs(positionDifference.X) > SystemParameters.MinimumHorizontalDragDistance
                && Math.Abs(positionDifference.Y) > SystemParameters.MinimumVerticalDragDistance)
            {
                TreeView treeView = sender as TreeView;
                TreeViewItem treeViewItem = FindAncestor<TreeViewItem>((DependencyObject)e.OriginalSource);

                Trace.WriteLine(e.OriginalSource.GetType());
                Trace.WriteLine(treeView.ItemContainerGenerator.ItemFromContainer(treeViewItem).GetType());
                ComosTreeViewNode comosTreeViewNode = treeView.ItemContainerGenerator.ItemFromContainer(treeViewItem) as ComosTreeViewNode;

                DataObject dragData = new DataObject("ComosTreeViewNode", comosTreeViewNode); 
                DragDrop.DoDragDrop(treeViewItem, dragData, DragDropEffects.Move);
            }
        }

        private void CopyTreeView_Drop(object sender, DragEventArgs e)
        {
            if(e.Data.GetDataPresent("ComosTreeViewNode"))
            {
                ComosTreeViewNode comosTreeViewNode = e.Data.GetData("ComosTreeViewNode") as ComosTreeViewNode;
                TreeView treeView = sender as TreeView;
                treeView.Items.Add(comosTreeViewNode);
            }
        }

        /// <summary>
        /// Pomáhá vyhledávání ve stromu
        /// </summary>
        /// <param name="current"></param>
        /// <returns></returns>
        private static T FindAncestor<T> (DependencyObject current)
            where T : DependencyObject
        {
            do
            {
                if(current is T)
                    return (T) current;

                current = VisualTreeHelper.GetParent(current);
            }
            while(current != null);
            return null;
        }

        private void My_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            mouseClickPoint = e.GetPosition(null);
        }

        private void SelectTemplateButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ComosTreeViewNode templateNode = TemplateTreeView.SelectedItem as ComosTreeViewNode;
                ComosTreeViewNode currentNode = CopyTreeView.SelectedItem as ComosTreeViewNode;

                LayeredCopyManager.CopyLayeredTemplate(templateNode, currentNode, mainHandler);


                //if (copyItem == null)
                //{
                //    mainHandler.CurrentRootNodes.Add(templateItem);
                //}
                //else if (copyItem != null)
                //{
                //    copyItem.Children.Add(templateItem);
                //    templateItem.Parent = copyItem;
                //}
            }
            catch(Exception ex)
            {
                CMessageBox.Show(ex.Message);
            }
        }

        private void PrepareCopyButton_Click(object sender, RoutedEventArgs e)
        {
            Type type= typeof(GlobalCastings);
            foreach (var method in type.GetMethods())
            {
                var parameters = method.GetParameters();
                var parameterDescriptions = string.Join
                    (", ", method.GetParameters()
                                 .Select(x => x.ParameterType + " " + x.Name)
                                 .ToArray());

                Console.WriteLine("{0} {1} ({2})",
                                  method.ReturnType,
                                  method.Name,
                                  parameterDescriptions);
            }

            //ComosProjTreeV.ProjTreeVega navigator = project.Workset().Globals().Navigator;
            //IComosDCollection selectedCollection = navigator.GetCurrentTree().SelectedObjectCollection();
            //IComosDDevice device;

            //if (selectedCollection.Count() == 0)
            //{
            //    device = navigator.PLTObject as IComosDDevice;
            //    if (device != null)
            //        Trace.WriteLine(device.DetailClass.ToString());
            //}

            //for (int i = 1; i <= selectedCollection.Count(); i++)
            //{
            //    device = selectedCollection.Item(i) as IComosDDevice;
            //    if (device != null)
            //        Trace.WriteLine(device.DetailClass.ToString());
            //}

        }
    }
}
