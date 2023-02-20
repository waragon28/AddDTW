using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Vistony.DTW.Win
{
    class FormDataEvent
    {
        public void SB1_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            switch (BusinessObjectInfo.FormTypeEx)
            {
                case "DispatchRoute":
                   
                    break;
            }
        }
    }
}
