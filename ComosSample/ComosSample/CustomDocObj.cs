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

namespace ComosSample
{
    internal class CustomDocObj
    {
        public void Test(IComosDDocument doc)
        {
            IComosDCollection docBackPointers = doc.BackPointerDocObjs();
            IComosDDocObj BpDocObj;
            IComosDDocument BpDoc;

            for(int i = 1; i <= docBackPointers.Count(); i++)
            {
                BpDocObj = docBackPointers.Item(i) as IComosDDocObj;
                BpDoc = BpDocObj.owner() as IComosDDocument;

                if(BpDoc != null && BpDocObj != null)
                {
                    Trace.WriteLine(BpDocObj.FullName());
                    Trace.WriteLine(BpDoc.FullName());
                    Trace.WriteLine(" ");
                }
            }
        }
    }
}
