using System;
using System.Collections.Generic;
using System.Text;
using SAPbouiCOM.Framework;
using Forxap.Framework.Constants;
using Forxap.Framework.UI;


namespace Vistony.DTW.Win
{
    public class SB1_FormMenuEvent
    {

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public void SB1_Application_FormMenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                
                switch (Application.SBO_Application.Forms.ActiveForm.TypeEx)
                {


                    case "DispatchRoute":
                        {
                         //   DispatchRoute .FrmDispatchRoute_FormMenuEvent(ref pVal, out BubbleEvent);
                            //UltimaMilla.FrmDispatchRoute.FrmDispatchRoute_FormMenuEvent(ref pVal, out BubbleEvent);
                  
                            break;
                       }





                }

                }
            catch (Exception)
            {

                // TODO: crear un log de forma asincrona
                //Forxap.Framework.UI.Messages.ShowError(ex.ToString());
            }



        }// fin del metodo FormMenuEvent



    }// fin de la clase

}// fin del namespace
