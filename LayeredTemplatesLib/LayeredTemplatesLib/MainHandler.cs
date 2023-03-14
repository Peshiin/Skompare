using Comos.Global;
using Plt;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;

namespace LayeredTemplatesLib
{
    public class MainHandler : INotifyPropertyChanged
    {
        private ObservableCollection<ComosTreeViewNode> templateRootNodes { get; set; }
        public ObservableCollection<ComosTreeViewNode> TemplateRootNodes { get { return templateRootNodes; } }
        private ObservableCollection<ComosTreeViewNode> currentRootNodes { get; set; }
        public ObservableCollection<ComosTreeViewNode> CurrentRootNodes { get { return currentRootNodes; } }
        private ObservableCollection<ComosTreeViewNode> copyList { get; set; }
        public ObservableCollection<ComosTreeViewNode> CopyList { get { return copyList; }}
        public ObservableCollection<IComosDWorkingOverlay> TemplateProjects { get ; set; }
            = new ObservableCollection<IComosDWorkingOverlay>();
        public ObservableCollection<string> TemplateProjectNames { get; set; }
            = new ObservableCollection<string>();
        private IComosDWorkset workset;

        public event PropertyChangedEventHandler PropertyChanged;
        /// <summary>
        /// Hlášení změny ve vlastnotech třídy prezentační vrstvě.
        /// </summary>
        protected void NotifyPropertyChanged(string propertyName)
        {
            var handler = PropertyChanged;
            if (handler != null)
                handler(this, new PropertyChangedEventArgs(propertyName));
        }

        /// <summary>
        /// Konstruktor třídy
        /// </summary>
        /// <exception cref="Exception"></exception>
        public MainHandler()
        {
            workset = Comos.Global.AppGlobal.Workset;
            if (workset == null)
                throw new Exception("Nenalezena databáze");

            templateRootNodes = new ObservableCollection<ComosTreeViewNode>();
            currentRootNodes = new ObservableCollection<ComosTreeViewNode>();
            copyList = new ObservableCollection<ComosTreeViewNode>();
            GetProjectTreeView(currentRootNodes, workset.GetCurrentProject());

            TemplateProjects = GetTemplateProjects();

            foreach(IComosDWorkingOverlay overlay in TemplateProjects)
                TemplateProjectNames.Add(overlay.FullName());
        }

        /// <summary>
        /// Vrátí kolekci pracovních vrstev template projektu
        /// </summary>
        /// <returns></returns>
        private ObservableCollection<IComosDWorkingOverlay> GetTemplateProjects()
        {
            ObservableCollection<IComosDWorkingOverlay> getTemplateProjects
                = new ObservableCollection<IComosDWorkingOverlay>();
            try
            {
                IComosDWorkset workset = this.workset;

                IComosDOwnCollection projects = workset.GetAllProjects();
                if (projects == null)
                    throw new Exception("Nenalezeny žádné projekty");

                IComosDProject project;
                IComosDWorkingOverlay workingLayer;
                IComosDCollection workingLayers;

                for (int i = 1; i <= projects.Count(); i++)
                {
                    project = projects.Item(i) as IComosDProject;
                    if (project != null && project.Type == "V")
                    {
                        workingLayers = project.WorkingOverlays();
                        if (workingLayers == null)
                        {
                            throw new Exception("Nenalezeny vrstvy projektu");
                        }
                        else
                        {
                            for (int j = 1; j <= workingLayers.Count(); j++)
                            {
                                workingLayer = workingLayers.Item(j) as IComosDWorkingOverlay;
                                getTemplateProjects.Add(workingLayer);
                            }
                        }
                    }
                }

                return getTemplateProjects;
            }

            catch (Exception ex)
            {
                CMessageBox.Show(ex.Message);
                return null;
            }
        }

        /// <summary>
        /// Ukládá strukturu projektu do TreeView struktury
        /// </summary>
        /// <param name="project"></param>
        public void GetProjectTreeView (ObservableCollection<ComosTreeViewNode>rootNodes, IComosDProject project)
        {

            ComosTreeViewNode root = new ComosTreeViewNode(null)
            {
                Description = project.Name,
                ComosObject = project as object,
                Parent = null
            };

            ComosTreeViewNode node;
            IComosDCollection projectDevices = project.Devices();
            IComosDDevice comosDevice;

            for(int i = 1; i <= projectDevices.Count(); i++)
            {
                comosDevice = projectDevices.Item(i) as IComosDDevice;
                node = new ComosTreeViewNode(root)
                {
                    Description = comosDevice.FullName(),
                    ComosObject = comosDevice as object,
                    Parent = root
                };
                root.Children.Add(node);
                GetComosSubdevicesToTreeNode(node);
            }

            rootNodes.Add(root);
            NotifyPropertyChanged(nameof(rootNodes));
        }

        /// <summary>
        /// Zapisuje děti nódu TreeView struktury
        /// </summary>
        /// <param name="treeViewOwner"></param>
        private void GetComosSubdevicesToTreeNode(ComosTreeViewNode treeViewOwner)
        {
            IComosDDevice comosOwner = treeViewOwner.ComosObject as IComosDDevice;
            ComosTreeViewNode node;
            IComosDCollection ownerDevices = comosOwner.Devices();
            IComosDDevice comosDevice;
            IComosDCollection ownerDocuments = comosOwner.Documents();
            IComosDDocument comosDocument;

            for(int i = 1; i <= ownerDevices.Count(); i++)
            {
                comosDevice = ownerDevices.Item(i);
                node = new ComosTreeViewNode(treeViewOwner)
                {
                    Description = comosDevice.FullName(),
                    ComosObject = comosDevice as object,
                    Parent = treeViewOwner
                };
                treeViewOwner.Children.Add(node);
                GetComosSubdevicesToTreeNode(node);
            }

            for (int i = 1; i <= ownerDocuments.Count(); i++)
            {
                comosDocument = ownerDocuments.Item(i);
                node = new ComosTreeViewNode(treeViewOwner)
                {
                    Description = comosDocument.FullName(),
                    ComosObject= comosDocument as object,
                    Parent = treeViewOwner
                };
                treeViewOwner.Children.Add(node);
            }
        }

    }
}
