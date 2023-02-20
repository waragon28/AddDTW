using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.Drawing;
using SAPbobsCOM;
using Forxap.Framework.Extensions;

using Forxap.Framework.UI;
using Vistony.DTW.Constans;
using System.Data;
using Vistony.DTW.Win;

namespace Vistony.DTW.Win
{
    public class Utils
    {
        private static object sb1message;
        public static AddonMessageInfo addonMessageInfo = new AddonMessageInfo();
      
        /// <summary>
        /// /
        /// </summary>
        /// <param name="oComboBox"></param>
        /// <param name="tableName"></param>
        public static void LoadMiscelaneo(ref SAPbouiCOM.ComboBox oComboBox, string tableName)
        {

            Dictionary<string, string> listObject;
            if (oComboBox != null)
            {

                listObject = Forxap.Framework.DI.Sb1Users.GetListFromUDT(tableName);
                foreach (var item in listObject)
                {
                    oComboBox.ValidValues.Add(item.Key, item.Value);
                }


            }
        }

        /// Carga un combobox dentro de una grilla 
        /// </summary>
        /// <param name="oComboBox"></param>
        public static void LoadMiscelaneo(ref SAPbouiCOM.Column oColumn, string tableName)
        {
            Dictionary<string, string> listObject;
            if (oColumn != null)
            {
                if (oColumn.Type == SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                {


                    listObject = Forxap.Framework.DI.Sb1Users.GetListFromUDT(tableName);
                    foreach (var item in listObject)
                    {
                        oColumn.ValidValues.Add(item.Key, item.Value);
                    }
                }
            }

        }



       
        public static bool findDataTable(SAPbouiCOM.DataTable datatable, string numeroPedido, string lineaPedido)
        {
            bool ret = false;

            string nroPedido = string.Empty;
            string linea = string.Empty;

            for (int i = 0; i < datatable.Rows.Count; i++)
            {
                nroPedido = datatable.GetValue("DocEntry", i).ToString();
                linea = datatable.GetValue("Lin", i).ToString();

                if ((numeroPedido == nroPedido) && (lineaPedido == linea))
                {
                    ret = true;
                }
            }


            return ret;
        }




        public static bool InitConfig()
        {
            SAPbobsCOM.Recordset recordSet = null;
            string code = string.Empty;
            bool ret = false;


            recordSet = (Recordset)Sb1Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            if (recordSet == null)
                throw new NullReferenceException("No se pudo obtener el objeto Recordset");

            try
            {

                string strSQL = "''";

                strSQL = string.Format("Select  \"Code\" From \"@FXP_DEMO_OADM\" where \"Code\" = 'CONFIG'  ");

                recordSet.DoQuery(strSQL);


                code = recordSet.Fields.Item("Code").Value.ToString();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            finally
            {
                if (recordSet != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(recordSet);
                    recordSet = null;
                    GC.Collect();
                }
            }

            // si no existe el registro de la configuración lo debo agregar
            if (code.Length == 0)
            {

                SAPbobsCOM.GeneralService oGeneralService = null;
                SAPbobsCOM.GeneralData oGeneralData = null;
                SAPbobsCOM.GeneralDataParams oGeneralParams = null;


                oGeneralService = Sb1Globals.oCompanyService.GetGeneralService("FXP_DEMO_OADM");


                oGeneralParams = ((SAPbobsCOM.GeneralDataParams)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)));


                oGeneralData = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);

                //     oGeneralData = oGeneralService.GetByParams(oGeneralParams);
                oGeneralData.SetProperty("Code", "CONFIG");



                oGeneralService.Add(oGeneralData);
                

                //Specify data for main UDO
                //

            }
            else
            {
                ret = true;
            }

            return ret;
        }

        public static string GetVehiculeCode(string placa, ref double? capacity , ref string brandName )
        {
            SAPbobsCOM.Recordset recordSet = null;
            string code = string.Empty;


            recordSet = (Recordset)Sb1Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            if (recordSet == null)
                throw new NullReferenceException("No se pudo obtener el objeto Recordset");

            try
            {

                string strSQL = "''";

                strSQL = string.Format("Select  \"Code\", \"U_SYP_VEPM\", \"U_SYP_VEMA\"  From \"@SYP_VEHICU\" where \"U_SYP_VEPL\" ='{0}'   ", placa);

                recordSet.DoQuery(strSQL);


                code = recordSet.Fields.Item("Code").Value.ToString();
                capacity = Convert.ToDouble( recordSet.Fields.Item("U_SYP_VEPM").Value);
                brandName = recordSet.Fields.Item("U_SYP_VEMA").Value.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }

            finally
            {
                if (recordSet != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(recordSet);
                    recordSet = null;
                    GC.Collect();
                }
            }

          
            return code;
        }


        public static DataTable GetCustomers(string ciudad)
        {
            SAPbobsCOM.Recordset recordSet = null;
            string code = string.Empty;


            recordSet = (Recordset)Sb1Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            if (recordSet == null)
                throw new NullReferenceException("No se pudo obtener el objeto Recordset");

            try
            {

                string strSQL = "''";

                strSQL = string.Format("CALL SP_VIS_POLY_TRAECLIENTES {'0'}", ciudad);

                recordSet.DoQuery(strSQL);

                
            }
            catch (Exception ex)
            {
                throw ex;
            }

            finally
            {
                if (recordSet != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(recordSet);
                    recordSet = null;
                    GC.Collect();
                }
            }


            // return code;
            return null;
        }


        public static string GetAyudandeName(string code)
        {
            SAPbobsCOM.Recordset recordSet = null;
            string name = string.Empty;

            recordSet = (Recordset)Sb1Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            if (recordSet == null)
                throw new NullReferenceException("No se pudo obtener el objeto Recordset");

            try
            {

                string strSQL = "''";

                strSQL = string.Format("Select  \"Name\" From \"@VIS_DIS_OAYD\" where \"Code\" ='{0}'   ", code);

                recordSet.DoQuery(strSQL);


                name = recordSet.Fields.Item("Name").Value.ToString();

            }
            catch (Exception ex)
            {
                throw ex;
            }

            finally
            {
                if (recordSet != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(recordSet);
                    recordSet = null;
                    GC.Collect();
                }
            }


            return name;
        }

        /// <summary>
        /// asigna el numero de registro dentro del Grid en la primera columna pone la númeracion
        /// verifica ´primero que existan registros
        /// </summary>
        /// <param name="oGrid"></param>
        /// <param name="oForm"></param>
        public static int AssignLineNro(ref SAPbouiCOM.Form oForm, ref SAPbouiCOM.Grid oGrid , ref double totalWeight)
        {
            //SAPbouiCOM.Form oForm = oApplication.Forms.ActiveForm;
           
            int rowIndex = 0;
           

            if (oGrid == null)
                throw new ArgumentNullException("oGrid");


            try
            {


                oForm.Freeze(true);




                if (!oGrid.DataTable.IsEmpty)
                {

                    for (rowIndex = 0; rowIndex <= oGrid.DataTable.Rows.Count - 1; rowIndex++)
                    {
                        oGrid.Columns.Item("RowsHeader").TitleObject.Caption = "#";
                        oGrid.RowHeaders.SetText(rowIndex , (rowIndex + 1 ).ToString());
                        oGrid.DataTable.SetValue("Marca", rowIndex, "Y");
                        totalWeight += oGrid.DataTable.GetDouble("Peso", rowIndex);
                    }
                }

                
            }
            catch (Exception ex)
            {

                Forxap.Framework.UI.Sb1Messages.ShowError(ex.ToString());
            }

            finally
            {
                oForm.Freeze(false);
            }
            // retorna el número de documentos seleccionados
            return rowIndex ;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="oForm"></param>
        /// <param name="oGrid"></param>
        /// <param name="totalWeight"></param>
        /// <param name="rowsSelected"></param>
        /// <returns></returns>
        public static int CheckRowsEx( SAPbouiCOM.Form oForm,  SAPbouiCOM.Grid oGrid, ref double totalWeight, ref int rowsSelected )
        {
            
            int rowIndex = 0;
            string newValue = "N";
            string oldValue = "Y";
            string consolidado = string.Empty;
            if (oGrid == null)
                throw new ArgumentNullException("oGrid");


            try
            {


                oForm.Freeze(true);




                if (!oGrid.DataTable.IsEmpty)
                {

                    for (rowIndex = 0; rowIndex <= oGrid.DataTable.Rows.Count - 1; rowIndex++)
                    {
                        // obtengo el valor si esta checkeado o no 
                        oldValue = oGrid.DataTable.GetString("Marca", rowIndex);

                        if (oldValue == "Y")
                        {
                            newValue = "N";
                            // obtengo el peso total DEL PESO del documento y le resto al total
                            totalWeight -= oGrid.DataTable.GetDouble("Peso", rowIndex);
                            // resto un row seleccionado al total de seleccionadps
                            rowsSelected -= 1;
                        }
                        else
                        {
                            
                                consolidado = oGrid.DataTable.GetValue("Consolidado", oGrid.GetDataTableRowIndex(rowIndex)).ToString();

                            // si no esta consolidado
                            if (consolidado.Length == 0)
                            {
                                newValue = "Y";
                                // obtengo el peso total del peso del  documento y le sumo al total
                                totalWeight += oGrid.DataTable.GetDouble("Peso", rowIndex);
                                // sumo un row seleccionado al total de seleccionadps
                                rowsSelected += 1;

                            }
                            else
                            {
                               // Sb1Messages.ShowError(AddonMessageInfo.Message208);
                            }
                        }

                        oGrid.DataTable.SetValue("Marca", rowIndex, newValue);

                        

                      
                    }

                    totalWeight = Math.Round(totalWeight, 2);
                }


            }
            catch (Exception ex)
            {

                Forxap.Framework.UI.Sb1Messages.ShowError(ex.ToString());
            }

            finally
            {
                oForm.Freeze(false);
            }

            return rowIndex;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="oForm"></param>
        /// <param name="oGrid"></param>
        /// <param name="totalWeight"></param>
        /// <param name="rowsSelected"></param>
        /// <returns></returns>
        public static int CheckRows(SAPbouiCOM.Form oForm, SAPbouiCOM.Grid oGrid, ref double totalWeight, ref int rowsSelected)
        {

            int rowIndex = 0;
            string newValue = "N";
            string oldValue = "Y";
            string consolidado = string.Empty;
            if (oGrid == null)
                throw new ArgumentNullException("oGrid");


            try
            {


                oForm.Freeze(true);




                if (!oGrid.DataTable.IsEmpty)
                {

                    for (rowIndex = 0; rowIndex <= oGrid.DataTable.Rows.Count - 1; rowIndex++)
                    {
                        // obtengo el valor si esta checkeado o no 
                        oldValue = oGrid.DataTable.GetString("Marca", rowIndex);

                        if (oldValue == "Y")
                        {
                            newValue = "N";
                            // obtengo el peso total DEL PESO del documento y le resto al total
                            totalWeight -= oGrid.DataTable.GetDouble("Peso", rowIndex);
                            // resto un row seleccionado al total de seleccionadps
                            rowsSelected -= 1;
                        }
                        else
                        {

                                                    
                                newValue = "Y";
                                // obtengo el peso total del peso del  documento y le sumo al total
                                totalWeight += oGrid.DataTable.GetDouble("Peso", rowIndex);
                                // sumo un row seleccionado al total de seleccionadps
                                rowsSelected += 1;

                          
                        }

                        oGrid.DataTable.SetValue("Marca", rowIndex, newValue);




                    }

                    totalWeight = Math.Round(totalWeight, 2);
                }


            }
            catch (Exception ex)
            {

                Forxap.Framework.UI.Sb1Messages.ShowError(ex.ToString());
            }

            finally
            {
                oForm.Freeze(false);
            }

            return rowIndex;
        }
        public static int CheckRowsEstadoDespacho(SAPbouiCOM.Form oForm, SAPbouiCOM.Grid oGrid,ref int rowsSelected)
        {

            int rowIndex = 0;
            string newValue = "N";
            string oldValue = "Y";
            string consolidado = string.Empty;
            if (oGrid == null)
                throw new ArgumentNullException("oGrid");


            try
            {


                oForm.Freeze(true);




                if (!oGrid.DataTable.IsEmpty)
                {

                    for (rowIndex = 0; rowIndex <= oGrid.DataTable.Rows.Count - 1; rowIndex++)
                    {
                        // obtengo el valor si esta checkeado o no 
                        oldValue = oGrid.DataTable.GetString("Marcar", rowIndex);

                        if (oldValue == "Y")
                        {
                            newValue = "N";
                            // resto un row seleccionado al total de seleccionadps
                            rowsSelected -= 1;
                        }
                        else
                        {


                            newValue = "Y";
                            // obtengo el peso total del peso del  documento y le sumo al total
                            // sumo un row seleccionado al total de seleccionadps
                            rowsSelected += 1;


                        }

                        oGrid.DataTable.SetValue("Marcar", rowIndex, newValue);
                        
                    }
                    
                }


            }
            catch (Exception ex)
            {

                Forxap.Framework.UI.Sb1Messages.ShowError(ex.ToString());
            }

            finally
            {
                oForm.Freeze(false);
            }

            return rowIndex;
        }


        


    }// fin de la clase

}// fin del namespace
