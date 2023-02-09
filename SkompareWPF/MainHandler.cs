using SkompareWPF.Components;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;

namespace SkompareWPF
{
    internal class MainHandler
    {
        public Excel.Application XlApp{ get; private set; } = new Excel.Application();
        public XlFile OldFile { get; private set; }
        public XlFile NewFile { get; private set; }

        public MainHandler(OpenFileControl oldControl, OpenFileControl newControl)
        {
            OldFile = new XlFile(oldControl, XlApp);
            NewFile = new XlFile(newControl, XlApp);
        }
    }
}
