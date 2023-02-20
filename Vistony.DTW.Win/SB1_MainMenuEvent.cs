using System;
using System.Collections.Generic;
using System.Text;

using SAPbouiCOM.Framework;

using Forxap.Framework.Extensions;
using Forxap.Framework.Constants;
using Forxap.Framework.UI;

using Vistony.DTW.Constans;
using Vistony.DTW.Win.Asistentes;

namespace Vistony.DTW.Win
{
    public class SB1_MainMenuEvent
    {

        static SAPbouiCOM.Form oForm;
        private string menuCaption = string.Empty;
        private SAPbouiCOM.Form activeForm = null;


        /// <summary>
        /// Capturo los eventos del Menu del AddOn Iker One

        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public void SB1_Application_MainMenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (pVal.BeforeAction)
                {
                    switch (pVal.MenuUID)
                    {
                        case AddonMenuItem.AddonMainMenu:
                            {
                                OnShowDTW();
                            }
                            break;

                    } // fin dl switch

                }// fin del IF


            }

            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
        }/// fin del  metodo




        // integración de SelesForce con SAP
        private void OnShowDTW()
        {
            try
            {
                DataTransferWorkbench form  = new DataTransferWorkbench();
                form.Show();
            }
            catch (Exception ex)
            {
                
                Forxap.Framework.UI.Sb1Messages.ShowError(ex.ToString());
            }
        }
        private void OnShowConfigForm(string tableName)
        {
          
            SAPbouiCOM.Button oButton = null;
            SAPbouiCOM.Matrix oMatrix = null;
            //SAPbouiCOM.MenuItem oMenuItem = null;
            //SAPbouiCOM.Menus oMenu = null;
            string menuID = string.Empty;
            string menuCaption = string.Empty;
            int index = 0;

            try
            {


                //obtiene el ID del menu que encontra a traves del nombre de la tablas
                menuID = Forxap.Framework.UI.Menu.GetMenuID(tableName, SboMenuItem.WindowsDefinedUser);

                
                Application.SBO_Application.ActivateMenuItem(menuID);

                menuCaption = Application.SBO_Application.Menus.Item(menuID).String;




                index = menuCaption.IndexOf("-", 0);



                activeForm = Application.SBO_Application.Forms.ActiveForm;
                activeForm.Freeze(true);

                activeForm.Title = menuCaption.Substring(index + 2);

                // SAPbouiCOM.Item oItem = null;
                // SAPbouiCOM.StaticText lblInfo = null;



                oButton = activeForm.GetButton("2");
                oMatrix = activeForm.GetMatrix("3");

                oMatrix.Columns.Item(3).TitleObject.Caption = "Código";
                oMatrix.Columns.Item(4).TitleObject.Caption = "Descripción";


                //    oMatrix.CommonSetting.MergeCell(1,3,true); 


                //oItem = activeForm.Items.Add("lblInfo1", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                //oItem.Left = oButton.Item.Left + 75;
                //oItem.Width = oMatrix.Item.Width;
                //oItem.Top = oButton.Item.Top - oButton.Item.Height;
                //oItem.Height = oButton.Item.Height;
                //oItem.Enabled = true;
                //lblInfo = ((SAPbouiCOM.StaticText)(oItem.Specific));
                //lblInfo.Caption = "Height : " + oItem.Height.ToString() + "Top : " +  oItem.Top.ToString();



                activeForm.Freeze(false);





            }
            catch (Exception ex)
            {
                Forxap.Framework.UI.Sb1Messages.ShowError(ex.ToString());
            }
            finally
            {

                // libero recursos

                if (activeForm != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(activeForm);
                    activeForm = null;
                    GC.Collect();
                }

                if (oMatrix != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMatrix);
                    oMatrix = null;
                    GC.Collect();
                }

            }

        }
        private void onShowEntrega()
        {
            try
            {

            }
            catch (Exception ex)
            {
                Forxap.Framework.UI.Sb1Messages.ShowError(ex.ToString());
            }
        }
        private void OnShowWzdProgramacion()
        {
            try
            {

               
            }
            catch (Exception ex)
            {
                Forxap.Framework.UI.Sb1Messages.ShowError(ex.ToString());
            }
        }
        private void onShowAsignaDespacho()
        {
            try
            {
               

            }
            catch (Exception ex)
            {
                Forxap.Framework.UI.Sb1Messages.ShowError(ex.ToString());
            }
        }
        private void onShowAsignaDespachoManual()
        {
            try
            {
              

            }
            catch (Exception ex)
            {
                Forxap.Framework.UI.Sb1Messages.ShowError(ex.ToString());
            }
        }
        private void onShowEstadoDespacho()
        {
            try
            {
             

            }
            catch (Exception ex)
            {
                Forxap.Framework.UI.Sb1Messages.ShowError(ex.ToString());
            }
        }
        private void OnShowSeguimiento()
        {
            try
            {

               

            }
            catch (Exception ex)
            {
                Forxap.Framework.UI.Sb1Messages.ShowError(ex.ToString());
            }
        }
        private void OnShowSeguimientoSectorista()
        {
            try
            {

                

            }
            catch (Exception ex)
            {
                Forxap.Framework.UI.Sb1Messages.ShowError(ex.ToString());
            }
        }
        private void OnShowActualizarUbigeo()
        {
            try
            {

                

            }
            catch (Exception ex)
            {
                Forxap.Framework.UI.Sb1Messages.ShowError(ex.ToString());
            }
        }
        private void OnShowChoferes()
        {
            try
            {
                Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "CONDUC", string.Empty);
                

            }
            catch (Exception ex)
            {
                Forxap.Framework.UI.Sb1Messages.ShowError(ex.ToString());
            }
        }

        private void OnShowAyudantes()
        {
            try
            {
                Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "OAYD", string.Empty);


            }
            catch (Exception ex)
            {
                Forxap.Framework.UI.Sb1Messages.ShowError(ex.ToString());
            }
        }
        private void OnShowActualizarIMEI()
        {
            try
            {
                // Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "VIS_DIS_ODRT", string.Empty);

                
            }
            catch (Exception ex)
            {
                Forxap.Framework.UI.Sb1Messages.ShowError(ex.ToString());
            }
        }
        private void OnShowRutaDespacho()
        {
            try
            {
                // Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "VIS_DIS_ODRT", string.Empty);

                
            }
            catch (Exception ex)
            {
                Forxap.Framework.UI.Sb1Messages.ShowError(ex.ToString());
            }
        }
        private void OnShowAndenes()
        {
            SAPbouiCOM.Matrix oMatrix = null;
            try
            {


                Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "OAND", string.Empty);


                activeForm = Application.SBO_Application.Forms.ActiveForm;

                activeForm.Freeze(true);
                oMatrix = activeForm.GetMatrix("3");

                oMatrix.Columns.Item(3).TitleObject.Caption = "Código x";
                oMatrix.Columns.Item(4).TitleObject.Caption = "Descripción x";
                oMatrix.Columns.Item(5).TitleObject.Caption = "Sucursal x";

                //        oColumn = (SAPbouiCOM.Column)oMatrix.Columns.Item("4");

                //  oCombobox = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item(5);
                //  oCombobox = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item(5).Cells.Item(oMatrix.VisualRowCount).Specific;

                //  _cmbExpDate = (SAPbouiCOM.ComboBox)oIMatrix.Columns.Item("iV_15")
                //  oColumn.Type = SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX;

                activeForm.Freeze(false);

            }
            catch (Exception ex)
            {
                Forxap.Framework.UI.Sb1Messages.ShowError(ex.ToString());
            }
        }

        private void OnShowChoferesUbigeo(string menuID)
        {
            try
            {
          

                Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "VIST_UBIGEOCHOFER", string.Empty);

                
                Application.SBO_Application.Forms.ActiveForm.Title = Application.SBO_Application.Menus.Item(menuID).String;
                Application.SBO_Application.Forms.ActiveForm.GetMatrix("3").Columns.Item(3).TitleObject.Caption = "Código";

  
            }
            catch (Exception ex)
            {
                Forxap.Framework.UI.Sb1Messages.ShowError(ex.ToString());
            }
        }
        

        private void OnShowVehiculos()
        {
            try
            {
                Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "VEHICU", string.Empty);


            }
            catch (Exception ex)
            {
                Forxap.Framework.UI.Sb1Messages.ShowError(ex.ToString());
            }
        }
        private void OnShowAsisAnulaDoc()
        {
            try
            {
                //wzdAnulacion wzd = new wzdAnulacion();

                //wzdConciliar wzd = new wzdConciliar();
          //      wzd.Show();
            }
            catch (Exception ex)
            {
                Forxap.Framework.UI.Sb1Messages.ShowError(ex.ToString());
            }
        }





    }// fin de la clase


}// fin del namespace
