using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Comos.Controls;
using Comos.Global;
using Comos.Global.AppControls;
using Plt;
using System.Diagnostics;
using System.ComponentModel;
using System.Collections.ObjectModel;

namespace LayeredTemplatesLib
{
    internal static class LayeredCopyManager
    {
        public static void test()
        {
            IComosDWorkset ws = AppGlobal.Workset;
            IAppControls ctrls = Comos.Global.AppGlobal.AppControls;
            IAppControl ctrl = ctrls[1].Control as IAppControl;
            IAppControlOptions ctrlOpts = ctrl.Options as IAppControlOptions;
            IAppControlAppearance ctrlAppear = ctrlOpts.Appearance as IAppControlAppearance;
            AppearanceType appearType = ctrlAppear.AppearanceType;
            IComosDSpecification spec;
            IComosDDocObj docObj;
        }
        public static void CopyLayeredTemplate(ComosTreeViewNode copyNode, ComosTreeViewNode targetNode, MainHandler mainHandler)
        {
            if (!(copyNode.ComosObject is IComosDDocument))
                throw new Exception("Jako template musí být vybrán dokument");
            if (!(targetNode.ComosObject is IComosDDevice))
                throw new Exception("Jako cílový objekt vyber device");

            IComosDWorkset workset = Comos.Global.AppGlobal.Workset;
            IComosDCopyManager comosCopyManager = workset.GetCopyManager();
            IComosDProject currentProject = workset.GetCurrentProject();
            var copyComosObject = (IComosDDocument)copyNode.ComosObject;
            var targetComosObject = (IComosDDevice)targetNode.ComosObject;

            comosCopyManager.DestinationProject = currentProject;
            comosCopyManager.DestinationObject = targetComosObject;
            comosCopyManager.SourceObjects().Add(copyNode.ComosObject);
            comosCopyManager.CollectReferences = false;
            comosCopyManager.Mode = 0;
            comosCopyManager.DocObjCopyMode = 0;

            workset.CrossProjectCopy(comosCopyManager);
            IComosDCollection newObjects = comosCopyManager.NewObjects();
            IComosDCollection newRootObjects = comosCopyManager.newrootobjects();
            IComosDCollection allNewDevices = GetAllDevices(newRootObjects);
            targetComosObject.Paste(newObjects, null, false);            

            ComosTreeViewNode node;
            IComosDDevice device;

            for (int i = 1; i <= newObjects.Count(); i++)
            {
                node = new ComosTreeViewNode(null);
                node.ComosObject = newObjects.Item(i);
                mainHandler.CopyList.Add(node);
                Trace.WriteLine(node.ToString());
            }

            for (int i = 1; i <= allNewDevices.Count(); i++)
            {
                device = allNewDevices.Item(i) as IComosDDevice;

                if (device == null)
                {
                    continue;
                }
                else
                {
                    node = new ComosTreeViewNode(null);
                    node.ComosObject = allNewDevices.Item(i);
                    Trace.WriteLine("Device má jméno: " + ((IComosDDevice)node.ComosObject).FullName());
                }
                
                if (device.Class == "D" && device.owner().Class == "U")
                {
                    device = (IComosDDevice)node.ComosObject;
                    if (device.spec("Z00T00002.Z00A00005") != null)
                        node.PS = device.spec("Z00T00002.Z00A00005").value;
                    if (device.spec("Z00T00002.Z00A00010") != null)
                        node.TS = device.spec("Z00T00002.Z00A00010").value;
                    if (device.spec("Z00T00002.Z00A00404") != null)
                        node.DN = device.spec("Z00T00002.Z00A00404").value;

                    mainHandler.CopyList.Add(node);
                }

                if (device.Class == "P" && device.Description.Contains("Pipe section"))
                {
                    device = (IComosDDevice)node.ComosObject;
                    if (device.spec("Z00T00004.Z00A00005") != null)
                        node.PS = device.spec("Z00T00004.Z00A00005").value;
                    if (device.spec("Z00T00004.Z00A00010") != null)
                        node.TS = device.spec("Z00T00004.Z00A00010").value;
                    if (device.spec("Z00T00004.Z00A00404") != null)
                        node.DN = device.spec("Z00T00004.Z00A00404").value;

                    mainHandler.CopyList.Add(node);
                }

            }
        }

        /// <summary>
        /// Získá všechny devices z kolekce objektů. Pro kopírovaní pomocí copy managera,
        /// který vrací newRootObjects, jako strom.
        /// </summary>
        /// <param name="collection"></param>
        /// <returns></returns>
        private static IComosDCollection GetAllDevices(IComosDCollection collection)
        {
            IComosDDevice device;

            for(int i = 1; i <= collection.Count(); i++)
            {
                device = collection.Item(i) as IComosDDevice;
                if (device == null)
                    continue;
                collection.Add(device);
                if (device.Devices().Count() > 0)
                    GetSubDevices(collection, device);                    
            }
            return collection;
        }

        private static IComosDCollection GetSubDevices(IComosDCollection collection, IComosDDevice parentDevice)
        {
            if(collection == null || parentDevice == null)
                return null;

            IComosDDevice subDevice;
            for(int i = 1; i <= parentDevice.Devices().Count(); i++)
            {
                subDevice = parentDevice.Devices().Item(i) as IComosDDevice;

                if (subDevice == null)
                    continue;

                collection.Add(subDevice);

                if(subDevice.Devices().Count() > 0)
                    GetSubDevices(collection, subDevice);
            }
            return collection;
        }

        /// <summary>
        /// Vrátí všechny nódy uložené ve stromu, které reprezentují COMOS dokumenty
        /// </summary>
        /// <param name="tree"></param>
        /// <returns></returns>
        private static IEnumerable<ComosTreeViewNode> GetDocumentNodes(IEnumerable<ComosTreeViewNode> tree)
        {
            IEnumerable<ComosTreeViewNode> result = new List<ComosTreeViewNode>();

            foreach(ComosTreeViewNode node in tree)
            {
                if (node.ComosObject is IComosDDocument)
                    result.Append(node);
            }
            return result;
        }

        /// <summary>
        /// Získá nódy template projektu referované ve vybraných nódech kopírovaného stromu
        /// </summary>
        /// <param name="documentNode"></param>
        /// <returns></returns>
        //public static ObservableCollection<ComosTreeViewNode> GetReferencedNodes(ComosTreeViewNode documentNode)
        //{
        //    IEnumerable<ComosTreeViewNode> result = new List<ComosTreeViewNode>();

        //    IComosDDocument document = documentNode.ComosObject as IComosDDocument;
        //    IComosDDocObj backPointerDocObj;

        //    IComosDDocument templateDocument;
        //    IComosDCollection templateDocumentBackPointers;
            
        //    foreach(ComosTreeViewNode node in templateDocumentNodes)
        //    {
        //        templateDocument = node.ComosObject as IComosDDocument;
        //        templateDocumentBackPointers = templateDocument.BackPointerDocObjs();

        //        for (int j = 1; j <= templateDocumentBackPointers.Count(); j++)
        //        {
        //            backPointerDocObj = templateDocumentBackPointers.Item(j);
        //            if (backPointerDocObj.owner() == document)
        //            {
        //                Trace.WriteLine("Našla se reference: ");
        //                Trace.WriteLine("Main document: "
        //                    + document.Name
        //                    + ", subdocument: "
        //                    + backPointerDocObj.Name);
        //                Trace.WriteLine("");

        //                result.Append(node);
        //            }
        //        }
        //    }
        //    return result;
        //}
    }
}
