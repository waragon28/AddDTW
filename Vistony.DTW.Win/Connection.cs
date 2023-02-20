using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using SAPbouiCOM.Framework;
using Forxap.Framework;
using Forxap.Framework.UI;
using Vistony.DTW.Win;
using Vistony.DTW.Constans;

namespace Vistony.DTW.Win
{
    public class Sb1Connection
    {
        /// <summary>
        /// Conecta el AddOn a la instancia de SAP que se encuentra cargada (desarrollo) o en la instancia
        /// que lo invoca en Producción.
        /// </summary>
        /// <returns></returns>
        public static bool ConnectToSAP()
        {
   
            bool ret = true;
           

            try
            {
               // AddonMessageInfo addonMessageInfo = new AddonMessageInfo();


                Sb1Globals.oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
                Sb1Globals.oCompanyService = Sb1Globals.oCompany.GetCompanyService();
                Sb1Globals.DbServerType = Sb1Globals.oCompany.DbServerType.ToString();
                Sb1Globals.Path = System.Windows.Forms.Application.StartupPath;

                Sb1Initialize.SetFramework(Application.SBO_Application, Sb1Globals.oCompany);//inicializa el framework.
                Vistony.DTW.DAL.BaseDAL.SetDataAccesLayer(Application.SBO_Application, Sb1Globals.oCompany);//inicializa el DAL
                Forxap.Framework.ServiceLayer.Sb1Initialize.SetFramework(Application.SBO_Application, Sb1Globals.oCompany);
                
                Sb1Globals.UserSignature = Sb1Globals.oCompany.UserSignature;// obtiene el codigo del usuario logeado
                Sb1Globals.UserName = Sb1Globals.oCompany.UserName; // Nombre  del usuario logeado ejemplo plazarte
               // Sb1Globals.Sucursal = Utils.GetSucursal(Sb1Globals.UserName);
               // Sb1Globals.Departamento = Utils.GetDepartamento(Sb1Globals.UserName);
                Sb1Globals.CompanyName = Sb1Globals.oCompany.CompanyName; // Nombre de la compañia a la que esta logeado
                
                if (Sb1Globals.Idioma == "")
                {
                    // Sb1Globals.Idioma = Utils.GetIdiomaSAP(Sb1Globals.UserName);
                     Sb1Globals.Idioma = "Spanish";
                }
                Sb1MetaData.AddMenusEspanol(); // agrega el menu del addon
                #region Creación del Menu, Iconos, Tablas, Campos, UDOs,Permisos, Scripts

                /*if (Sb1Globals.Idioma == "English (United States)")
                {
                    Sb1MetaData.AddMenusEnglisUnitedStates(); // agrega el menu del addon
                }
                else if (Sb1Globals.Idioma == "French")
                {
                    Sb1MetaData.AddMenusFrench(); // agrega el menu del addon
                }*/

                Sb1Messages.ShowSuccess(string.Format(AddonMessageInfo.StartLoading), SAPbouiCOM.BoMessageTime.bmt_Short);

                Sb1MetaData.AddIcon();
                Sb1MetaData.AddUserProcedure();
                #endregion

                Sb1Messages.ShowSuccess(string.Format(AddonMessageInfo.FinishLoading), SAPbouiCOM.BoMessageTime.bmt_Short);

                ret = true;
            }
            catch (Exception ex)
            {
                //  if (Messages != null)
                Sb1Messages.ShowError(ex.Message);
            }



            return ret;

        }// fin del metodo connectToSAP

 



    }// fin de la clase

}// fin del namespace
