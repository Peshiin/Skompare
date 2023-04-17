using System;
using System.Collections.Generic;
using System.Linq;
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
using Comos.Controls;
using Comos.Global;
using Comos.Global.AppControls;
using ComosVBInterface;
using Plt;

namespace ComosSample
{
    /// <summary>
    /// Interakční logika pro MyCustomControl.xaml
    /// </summary>
    public partial class MyCustomControl: IComosControl
    {

        String strAttribute { get; set; }

        public MyCustomControl()
        {
            InitializeComponent();
            try
            {
            }
            catch(Exception ex)
            {
                Comos.Global.CMessageBox.Show(ex.Message);
            }
            
        }

        /// <summary>
        /// Loads DSPW process template project
        /// </summary>
        public void LoadTemplateProject()
        {
            try
            {
                IComosDWorkset workset = Comos.Global.AppGlobal.Workset;
                if (workset == null)
                    throw new Exception("Nenalezena databáze");

                IComosDOwnCollection projects = workset.GetAllProjects();
                if (projects == null)
                    throw new Exception("Nenalezeny žádné projekty");

                IComosDProject currentProject = workset.GetCurrentProject();
                if (currentProject == null)
                    throw new Exception("Není spuštěn žádný projekt");

                IComosDProject project = null;

                for (int i = 1; i <= projects.Count(); i++)
                {
                    project = projects.Item(i);
                    if (project != null && project.Name == "Template_DSPW")
                    {
                        IComosDCollection workingLayers = project.WorkingOverlays();
                        if (workingLayers == null)
                            throw new Exception("Nenalezeny vrstvy projektu") ;

                        IComosDWorkingOverlay workingLayer = workingLayers.Item(2);

                        Comos.Global.AppGlobal.ChangeProject(project);
                        project.let_CurrentWorkingOverlay(workingLayer);
                        Comos.Global.AppGlobal.ChangeProject(currentProject);

                        return;
                    }
                }

                throw new Exception("Nenalezen projekt \"Template_DSPW\"");
            }

            catch (Exception ex)
            {
                CMessageBox.Show(ex.Message);
            }            
        }

        IComosDWorkset IComosControl.Workset { get; set; }
        IComosDGeneralCollection IComosControl.Objects { get; set; }
        string IComosControl.Parameters { get; set; }
        IContainer IComosControl.ControlContainer { get; set; }

        void IComosControl.OnCanExecute(CanExecuteRoutedEventArgs e)
        {
            Console.WriteLine(e.ToString());
            Console.WriteLine("OnCanExecute()");
        }

        void IComosControl.OnExecuted(ExecutedRoutedEventArgs e)
        {
            Console.WriteLine(e.ToString());
            Console.WriteLine("OnExecuted()");
        }

        void IComosControl.OnPreviewExecuted(ExecutedRoutedEventArgs e)
        {
            Console.WriteLine(e.ToString());
            Console.WriteLine("OnPreviewExecuted()");
        }


        /// <summary>
        /// Vrací DSPW template projekt
        /// </summary>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public IComosDProject GetDspwTemplateProject()
        {
            IComosDWorkset workset = Comos.Global.AppGlobal.Workset;
            IComosDOwnCollection projects = workset.GetAllProjects();
            IComosDProject project;

            for (int i = 1; i <= projects.Count() + 1; i++)
            {
                project = projects.Item(i);
                if (project != null && project.Name == "Template_DSPW")
                {
                    return project;
                }

                if (i == projects.Count() + 1)
                    throw new Exception("Nenalezena template projekt");
            }
            return null;
        }

        private void GetProjectTreeView(TreeView treeView, IComosDProject project)
        {
            try
            {
                IComosDCollection templateDevicesCollection = null;

                IComosDDevice item;
                TreeViewItem treeViewItem;

                templateDevicesCollection = project.Devices();

                for (int i = 1; i <= templateDevicesCollection.Count(); i++)
                {
                    item = templateDevicesCollection.Item(i);

                    if (item != null)
                    {
                        treeViewItem = new TreeViewItem()
                        {
                            Header = item.Name + "  |  " + item.Description,
                            Tag = project.PathFullName(item),
                            TabIndex = i
                        };
                        treeView.Items.Add(treeViewItem);


                        WriteTreeViewDocuments(treeView, treeViewItem, item);
                        writeTreeViewSubitems(treeView, treeViewItem, item);
                    }
                }
            }
            catch (Exception ex)
            {
                CMessageBox.Show(ex.Message);
            }
        }

        void WriteTreeViewDocuments(TreeView treeView, TreeViewItem treeViewParent, IComosDDevice item)
        {
            IComosDDevice ownerDevice = item.owner() as IComosDDevice;
            IComosDProject ownerProject = item.owner() as IComosDProject;
            IComosDCollection templateDocumentsCollection;
            IComosDDocument document;
            TreeViewItem treeViewItem;

            if (ownerProject != null)
            {
                templateDocumentsCollection = ownerProject.Documents();
            }
            else if (ownerDevice != null)
            {
                templateDocumentsCollection = ownerDevice.Documents();
            }
            else
                throw new Exception("Owner není ani projekt ani device");


            for (int i = 1; i <= templateDocumentsCollection.Count(); i++)
            {
                document = templateDocumentsCollection.Item(i);
                if (document != null && document.owner() == item.owner())
                {
                    treeViewItem = new TreeViewItem()
                    {
                        Header = document.Name + "  |  " + document.Description,
                        Tag = item.PathFullName(document),
                        TabIndex = i
                    };
                    treeViewParent.Items.Add(treeViewItem);
                }
            }
        }

        /// <summary>
        /// Vyhledá objekt příslušného jména v zadané oblasti
        /// </summary>
        /// <param name="name"></param>
        /// <param name="searchArea"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        private object GetObjectByName(string name, IComosDCollection searchArea)
        {
            for(int i = 1; i < searchArea.Count()+1; i++)
            {
                if(searchArea.Item(i).Name == name)
                    return searchArea.Item(i);
                    
            }
            throw new Exception("Nenalezen objekt příslušného jména");
        }

        private void actionButton_Click(object sender, RoutedEventArgs e)
        {
            IComosDProject templateProject = GetDspwTemplateProject();
            IComosDCollection templateDocuments = templateProject.AllDocuments();
            IComosDDocument document = (IComosDDocument)GetObjectByName("DRAIN_1", templateDocuments);
            CustomDocObj customDocObj = new CustomDocObj();
            customDocObj.Test(document);

            IAppControls appCtrls = Comos.Global.AppGlobal.AppControls;
            IAppControl appCtrl = appCtrls[0];
            IAppControlOptions appCtrlOpts = appCtrl.Options;
            IAppControlAppearance ctrlAppear = appCtrlOpts.Appearance;

            //Vypsání projektů do treeView
            //GetProjectTreeView(templateTreeView, GetDspwTemplateProject());
            //GetProjectTreeView(currentTreeView, Comos.Global.AppGlobal.Workset.GetCurrentProject());

            //Hledání objektů na dokumentu
            //IComosDProject templateProject = getDSPWTemplateProject();
            //IComosDCollection templateDocuments = templateProject.AllDocuments();
            //IComosDDocument document = (IComosDDocument) GetObjectByName("MAW10_1", templateDocuments);
            //IComosDDevice documentOwner = document.owner();
            //IComosDCollection deviceCollection = documentOwner.ScanDevices("D*");
            //IComosDDevice device;
            //IComosDCollection deviceBpDocObjs;
            //IComosDDocument docObjOwner;
            //IComosDCollection docObjOwnerBPDocs;
            //TreeViewItem treeViewItem;

            //for (int i = 1; i < deviceCollection.Count() + 1; i++)
            //{
            //    device = deviceCollection.Item(i) as IComosDDevice;
            //    if (device != null)
            //    {
            //        deviceBpDocObjs = device.BackPointerDocObjs();
            //        docObjOwner = deviceBpDocObjs.Item(1).Owner;
            //        docObjOwnerBPDocs = docObjOwner.BackPointerDocObjs();

            //        treeViewItem = new TreeViewItem()
            //        {
            //            Header = deviceBpDocObjs.Count().ToString() + " | "
            //                    + deviceBpDocObjs.Item(1).Owner.Name + " | "
            //                    + deviceBpDocObjs.Item(1).Label + " || "
            //                    + docObjOwnerBPDocs.Count() + " | "
            //                    + docObjOwnerBPDocs.Item(1).Owner.Name + " | "
            //                    + deviceBpDocObjs.Item(1).Label
            //        };
            //        currentTreeView.Items.Add(treeViewItem);

            //    }

            //}
        }

        private void writeTreeViewSubitems(TreeView treeView, TreeViewItem treeViewParent, IComosDDevice comosOwner)
        {
            TreeViewItem item;
            IComosDDevice device;
            for (int i = 1; i < comosOwner.Devices().Count() + 1; i++)
            {
                device = comosOwner.Devices().Item(i);
                if (device != null && device.Class == "E")
                    return;

                if (device != null)
                {
                    item = new TreeViewItem()
                    {
                        Header = device.Name + "  |  " + device.Description,
                        Tag = comosOwner.PathFullName(device),
                        TabIndex = i,
                    };
                    treeViewParent.Items.Add(item);
                    WriteTreeViewDocuments(treeView, item, device);
                    writeTreeViewSubitems(treeView, item, device);
                }
            }
        }

        private void copyTemplateButton_Click(object sender, RoutedEventArgs e)
        {
            TreeViewItem selectedTemplateItem = (TreeViewItem)templateTreeView.SelectedItem;
            IComosDProject templateProject = GetDspwTemplateProject();
            IComosDDevice templateFolder = (IComosDDevice)templateProject.GetObjectByPathFullName(selectedTemplateItem.Tag.ToString());
            IComosDDevice templateDevice = templateFolder.Devices().Item(selectedTemplateItem.TabIndex);

            TreeViewItem selectedCurrentItem = (TreeViewItem)currentTreeView.SelectedItem;
            IComosDProject currentProject = Comos.Global.AppGlobal.Workset.GetCurrentProject();
            IComosDDevice currentFolder = currentProject.GetObjectByPathFullName(selectedCurrentItem.Tag.ToString()) as IComosDDevice;
            IComosDDevice currentDevice;
            if (currentFolder != null)
            {
                currentDevice = currentFolder.Devices().Item(selectedCurrentItem.TabIndex);
                copyTemplate(currentDevice, templateDevice);
            }
            else
            {
                currentDevice = currentProject.Devices().Item(selectedCurrentItem.TabIndex);
                copyTemplate(currentDevice, templateDevice);
            }
        }

        private void copyTemplate(IComosDDevice target,
                                    IComosDDevice template)
        {
            CMessageBox.Show(template.Description);

            template.CopyAll();
            target.Paste2(template);

        }

        private void copyDocumentButton_Click(object sender, RoutedEventArgs e)
        {

            TreeViewItem selectedTemplateItem = (TreeViewItem)templateTreeView.SelectedItem;
            IComosDProject templateProject = GetDspwTemplateProject();
            IComosDDevice templateFolder = (IComosDDevice)templateProject.GetObjectByPathFullName(selectedTemplateItem.Tag.ToString());
            IComosDDocument templateDocument = templateFolder.Documents().Item(selectedTemplateItem.TabIndex);

            TreeViewItem selectedCurrentItem = (TreeViewItem)currentTreeView.SelectedItem;
            IComosDProject currentProject = Comos.Global.AppGlobal.Workset.GetCurrentProject();
            IComosDDevice currentFolder = currentProject.GetObjectByPathFullName(selectedCurrentItem.Tag.ToString()) as IComosDDevice;
            IComosDDevice currentDevice;
            if (currentFolder != null)
            {
                currentDevice = currentFolder.Devices().Item(selectedCurrentItem.TabIndex);
                copyDocument(currentDevice, templateDocument);
            }
            else
            {
                currentDevice = currentProject.Devices().Item(selectedCurrentItem.TabIndex);
                copyDocument(currentDevice, templateDocument);
            }

        }
        private void copyDocument(IComosDDevice target,
                                    IComosDDocument template)
        {
            CMessageBox.Show(template.Description);

            template.CopyAll();
            target.Paste2(template);

        }
    }
}
