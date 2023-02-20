using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using System.Net;
using System.IO;
using Forxap.Framework.Extensions;
using Vistony.DTW.Constans;
using SAPbouiCOM;
using Forxap.Framework.UI;
using Forxap.Framework.Constants;
using System.Threading;
using System.Windows.Forms;
using ICSharpCode.SharpZipLib.Zip;
using Vistony.DTW.Win;
using ExcelDataReader;
using System.Data;

namespace Vistony.DTW.Win.Asistentes
{
    [FormAttribute("Vistony.DTW.Win.Asistentes.DataTransferWorkbench", "Asistentes/DataTransferWorkbench.b1f")]
    class DataTransferWorkbench : UserFormBase
    {
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.OptionBtn OptionBtn0;
        private SAPbouiCOM.OptionBtn OptionBtn1;
        private SAPbouiCOM.OptionBtn OptionBtn2;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.PictureBox PictureBox0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.StaticText StaticText5;
        private SAPbouiCOM.StaticText StaticText6;
        private SAPbouiCOM.StaticText StaticText7;
        private SAPbouiCOM.StaticText StaticText8;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.Button Button3;
        private SAPbouiCOM.Button Button4;
        private SAPbouiCOM.Button Button5;
        public SAPbouiCOM.Form oForm;
        private SAPbouiCOM.StaticText StaticText9;
        private SAPbouiCOM.StaticText StaticText10;
        private SAPbouiCOM.StaticText StaticText12;
        private SAPbouiCOM.OptionBtn OptionBtn3;
        private SAPbouiCOM.OptionBtn OptionBtn4;
        private SAPbouiCOM.OptionBtn OptionBtn5;
        public SAPbouiCOM.Matrix oMatrix;
        public SAPbouiCOM.Matrix Matrix0;
        public int PaneLevel { get; set; }
        public int PaneMax = 4;
        LabelsForms labelsForms = new LabelsForms();
        private void Anterior()
        {
            if (oForm.PaneLevel + 1 >= 2)
            {
                oForm.PaneLevel -= 1;
            }
            if (oForm.PaneLevel == 1)
            {

            }
            if (oForm.PaneLevel == 3)
            {
               
            }
        }
        
        public void Siguiente()
        {
            if (PaneLevel < PaneMax)
            {
                oForm.PaneLevel += 1;
            }
            if (oForm.PaneLevel==5)
            {
                addItem();
            }
            if (oForm.PaneLevel==2)
            {
                OptionBtn1.GroupWith("Item_1");
                OptionBtn2.GroupWith("Item_1");
                oForm.SetUserDataSource("UD_0", "Y");
                OptionBtn0.Selected = true;
                oForm.PaneLevel = 2;
            }
            if (oForm.PaneLevel == 3)
            {
                OptionBtn4.GroupWith("Item_22");
                OptionBtn5.GroupWith("Item_22");
                oForm.SetUserDataSource("UD_5", "Y");
                OptionBtn3.Selected = true;
                oForm.PaneLevel = 3;
            }
            if (oForm.PaneLevel == 4)
            {

                Button7.Item.Click();
            }
            if (oForm.PaneLevel == 6)
            {
                int Registros = Matrix0.RowCount;
                if (Registros==1)
                {
                    for (int oRow = 0; oRow < Matrix0.RowCount; oRow++)
                    {
                        SAPbouiCOM.DataTable udt = oForm.GetDataTable("DB_Arbol");
                        string archivo = Convert.ToString(udt.GetValue("Archivo", oRow));
                        FileStream fStream = File.Open(archivo, FileMode.Open, FileAccess.Read);
                        IExcelDataReader excelDataReader = ExcelReaderFactory.CreateOpenXmlReader(fStream);
                        DataSet result = excelDataReader.AsDataSet();
                        excelDataReader.Close();
                        /*using (var stream = File.Open(archivo, FileMode.Open, FileAccess.Read))
                        {
                            using (var reader = ExcelReaderFactory.CreateReader(stream))
                            {
                                var result = reader.AsDataSet();
                                // Ejemplos de acceso a datos
                                // Stream table = result.Tables[0];
                                //  DataRow row = table.Rows[0];
                                string cell = result.Tables[0].ToString();
                            }
                        }*/

                    }

                }
                else
                {

                }
            }
        }
        public DataTransferWorkbench()
        {
           

        }
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.OptionBtn0 = ((SAPbouiCOM.OptionBtn)(this.GetItem("Item_1").Specific));
            this.OptionBtn0.PressedAfter += new SAPbouiCOM._IOptionBtnEvents_PressedAfterEventHandler(this.OptionBtn0_PressedAfter);
            this.OptionBtn0.ClickAfter += new SAPbouiCOM._IOptionBtnEvents_ClickAfterEventHandler(this.OptionBtn0_ClickAfter);
            this.OptionBtn1 = ((SAPbouiCOM.OptionBtn)(this.GetItem("Item_2").Specific));
            this.OptionBtn1.ClickAfter += new SAPbouiCOM._IOptionBtnEvents_ClickAfterEventHandler(this.OptionBtn1_ClickAfter);
            this.OptionBtn2 = ((SAPbouiCOM.OptionBtn)(this.GetItem("Item_3").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_4").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_5").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_6").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_7").Specific));
            this.PictureBox0 = ((SAPbouiCOM.PictureBox)(this.GetItem("Item_8").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_9").Specific));
            this.Button1.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button1_ClickBefore);
            this.Button1.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button1_ClickAfter);
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_10").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_11").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_12").Specific));
            this.StaticText7 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_13").Specific));
            this.StaticText8 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_14").Specific));
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_15").Specific));
            this.Button3 = ((SAPbouiCOM.Button)(this.GetItem("Item_16").Specific));
            this.Button3.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button3_ClickAfter);
            this.Button4 = ((SAPbouiCOM.Button)(this.GetItem("Item_17").Specific));
            this.Button4.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button4_ClickAfter);
            this.Button5 = ((SAPbouiCOM.Button)(this.GetItem("Item_18").Specific));
            this.StaticText9 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.StaticText10 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_19").Specific));
            this.StaticText12 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_21").Specific));
            this.OptionBtn3 = ((SAPbouiCOM.OptionBtn)(this.GetItem("Item_22").Specific));
            this.OptionBtn4 = ((SAPbouiCOM.OptionBtn)(this.GetItem("Item_23").Specific));
            this.OptionBtn5 = ((SAPbouiCOM.OptionBtn)(this.GetItem("Item_24").Specific));
            this.StaticText11 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_27").Specific));
            this.Grid1 = ((SAPbouiCOM.Grid)(this.GetItem("Item_29").Specific));
            this.Grid1.ClickAfter += new SAPbouiCOM._IGridEvents_ClickAfterEventHandler(this.Grid1_ClickAfter);
            this.Button7 = ((SAPbouiCOM.Button)(this.GetItem("Item_30").Specific));
            this.Button7.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button7_ClickAfter);
            this.Button8 = ((SAPbouiCOM.Button)(this.GetItem("Item_31").Specific));
            this.Button8.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button8_ClickAfter);
            this.StaticText13 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_20").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_25").Specific));
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_28").Specific));
            this.Matrix0.ClickAfter += new SAPbouiCOM._IMatrixEvents_ClickAfterEventHandler(this.Matrix0_ClickAfter);
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_32").Specific));
            this.Button0.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button0_ClickAfter);
            this.Grid2 = ((SAPbouiCOM.Grid)(this.GetItem("Item_26").Specific));
            this.Grid4 = ((SAPbouiCOM.Grid)(this.GetItem("Item_34").Specific));
            this.Grid5 = ((SAPbouiCOM.Grid)(this.GetItem("Item_35").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }
        

        private void OnCustomInitialize()
        {
            oForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.Item(this.UIAPIRawForm.UniqueID);
            oForm.ScreenCenter();
            /*FORMULARIO 1*/
            oForm.Title = labelsForms.LabelsFormsTitulo(Sb1Globals.Idioma);
            StaticText2.Caption = labelsForms.LabelsFormsLabel00001(Sb1Globals.Idioma);
            StaticText3.Caption = labelsForms.LabelsFormsLabel00002(Sb1Globals.Idioma);
            StaticText8.Caption= labelsForms.LabelsFormsLabel00003(Sb1Globals.Idioma);
            StaticText4.Caption = labelsForms.LabelsFormsLabel00004(Sb1Globals.Idioma);
            StaticText5.Caption = labelsForms.LabelsFormsLabel00005(Sb1Globals.Idioma);
            StaticText6.Caption = labelsForms.LabelsFormsLabel00006(Sb1Globals.Idioma);
            StaticText7.Caption = labelsForms.LabelsFormsLabel00007(Sb1Globals.Idioma);

            Button5.Caption = labelsForms.LabelsFormsLabelFinalizar(Sb1Globals.Idioma);
            Button1.Caption = labelsForms.LabelsFormsLabelSiguiente(Sb1Globals.Idioma);
            Button4.Caption = labelsForms.LabelsFormsLabelAntes(Sb1Globals.Idioma);
            Button3.Caption = labelsForms.LabelsFormsLabelCancelar(Sb1Globals.Idioma);
            Button2.Caption = labelsForms.LabelsFormsLabelAyuda(Sb1Globals.Idioma);
           
            /*FORMULARIO 2*/
            StaticText0.Caption = labelsForms.LabelsFormsLabel00008(Sb1Globals.Idioma);
            StaticText1.Caption = labelsForms.LabelsFormsLabel00009(Sb1Globals.Idioma);
            OptionBtn0.Caption = labelsForms.LabelsFormsLabel00010(Sb1Globals.Idioma);
            OptionBtn1.Caption = labelsForms.LabelsFormsLabel00011(Sb1Globals.Idioma);
            OptionBtn2.Caption = labelsForms.LabelsFormsLabel00012(Sb1Globals.Idioma);

            /*FORMULARIO 3*/
            StaticText9.Caption = labelsForms.LabelsFormsLabel00013(Sb1Globals.Idioma);
            StaticText10.Caption = labelsForms.LabelsFormsLabel00014(Sb1Globals.Idioma);
            StaticText12.Caption = labelsForms.LabelsFormsLabel00015(Sb1Globals.Idioma);
            OptionBtn3.Caption = labelsForms.LabelsFormsLabel00016(Sb1Globals.Idioma);
            OptionBtn4.Caption = labelsForms.LabelsFormsLabel00017(Sb1Globals.Idioma);
            OptionBtn5.Caption = labelsForms.LabelsFormsLabel00018(Sb1Globals.Idioma);

            StaticText11.Caption = labelsForms.LabelsFormsLabel00019(Sb1Globals.Idioma);

            /*FORMULARIO 4*/
            StaticText13.Caption = labelsForms.LabelsFormsLabel000021(Sb1Globals.Idioma);
            Button7.Caption= labelsForms.LabelsFormsLabelDescargar(Sb1Globals.Idioma);
            Button8.Caption = labelsForms.LabelsFormsLabelComprimir(Sb1Globals.Idioma);

            OptionBtn1.GroupWith("Item_1");
            OptionBtn2.GroupWith("Item_1");
            oForm.SetUserDataSource("UD_0", "Y");
            OptionBtn0.Selected = true;


            OptionBtn4.GroupWith("Item_22");
            OptionBtn5.GroupWith("Item_22");
            oForm.SetUserDataSource("UD_5", "Y");
            OptionBtn3.Selected = true;

            StaticText2.SetBold();
            StaticText2.SetSize(15);

            StaticText0.SetBold();
            StaticText0.SetSize(15);

            StaticText9.SetBold();
            StaticText9.SetSize(15);

            StaticText8.SetBold();
            StaticText8.SetSize(15);

            StaticText11.SetBold();
            StaticText11.SetSize(15);

            Matrix0.AutoResizeColumns();
            Grid1.AutoResizeColumns();
            //oForm.Refresh();
        }


        private void Button3_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            oForm.Close();
        }

        private void Button1_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            Siguiente();
        }

        private void Button4_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            Anterior();
        }

        private void OptionBtn0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
           // throw new System.NotImplementedException();

        }

        private void OptionBtn1_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
        }

        private void Button0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (SelectMatrix>0)
            {
                Thread t = new Thread(() =>
                {
                    OpenFileDialog openFileDialog = new OpenFileDialog();
                    openFileDialog.Title = "Seleccione un archivo";
                    openFileDialog.Filter = "Archivo de texto | *.txt | Archivo CSV | *.csv";

                    DialogResult dr = openFileDialog.ShowDialog(new System.Windows.Forms.Form());
                    if (dr == DialogResult.OK)
                    {
                        string fileName = openFileDialog.FileName;
                        SAPbouiCOM.DataTable udt = oForm.GetDataTable("DB_Arbol");
                        udt.SetValue("Archivo", SelectMatrix-1, fileName);
                        Matrix0.LoadFromDataSource();
                        //SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(fileName);
                    }
                });          // Kick off a new thread
                t.IsBackground = true;
                t.SetApartmentState(ApartmentState.STA);
                t.Start();
            }
            else
            {
                Sb1Messages.ShowError("Debe Seleccionar un Registro");
            }
        }
        private void addItem()
        {
            try
            {
                string TipoData = "";
                if (OptionBtn0.Selected)
                {
                    TipoData = "Datos de configuración";
                }
                else if (OptionBtn1.Selected)
                {
                    TipoData = "Master Data";
                }
                else if (OptionBtn2.Selected)
                {
                    TipoData = "Datos de configuración";
                }

                /*EJECUTAR EL PROCEDIMIENTO ALMACENADO*/
                SAPbouiCOM.DataTable exp;
                exp = oForm.DataSources.DataTables.Item("Execute_DT");
                string Query = string.Format(AddonMessageInfo.Query_List_Table_Obj_DTW, TipoData, Obj_DTW);
                exp.ExecuteQuery(Query);
                /*FIN DE EJECUCION DE PROCEDIMIENTO ALMACENADO*/

                SAPbouiCOM.DataTable udt = oForm.GetDataTable("DB_Arbol");
                oMatrix = oForm.GetMatrix("Item_28");
                SAPbouiCOM.Columns oColumns;
                oColumns = oMatrix.Columns;
                SAPbouiCOM.Column oColumn;
                var colItems = udt.Columns;
                if (udt.Columns.Count == 0)
                {
                    colItems.Add("Tabla", BoFieldsType.ft_AlphaNumeric);
                    colItems.Add("Archivo", BoFieldsType.ft_AlphaNumeric);
                }
                int a = udt.Rows.Count;
                if (oMatrix.RowCount > 0)
                    a = udt.Rows.Count;
                for (int oRow = 0; oRow < exp.Rows.Count; oRow++)
                {
                    udt.Rows.Add();
                    udt.SetValue("Tabla", oRow, exp.GetString("Tabla", oRow));
                    udt.SetValue("Archivo", oRow, exp.GetString("Archivo", oRow));
                }

                oMatrix.Columns.Item("Col_0").DataBind.Bind("DB_Arbol", "Tabla");
                oMatrix.Columns.Item("Col_1").DataBind.Bind("DB_Arbol", "Archivo");
                
                oColumn = oColumns.Item("Col_0");
                oMatrix.LoadFromDataSourceEx();
                oMatrix.AutoResizeColumns();



            }
            catch (Exception ex)
            {

                throw;
            }
        }

        private void Matrix0_CollapsePressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
          //  Matrix0.CollapsePressedAfter();
        }

        private Grid Grid0;

        public SAPbouiCOM.DataTable Dow(string Query, SAPbouiCOM.DataTable oDT)
        {
            try
            {
                string StrHANA = Query;
                oDT.ExecuteQuery(StrHANA);
                return oDT;
            }
            catch (Exception ex)
            {
                ex.StackTrace.ToString();
                return null;
            }
        }

        private StaticText StaticText11;
        private Grid Grid1;
        private SAPbouiCOM.Button Button7;
        string Obj_DTW = "";
        private void Button7_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            oForm.Freeze(true);
            /*TIPO DE DATA*/
            string TipoData = "";
            if (OptionBtn0.Selected)
            {
                TipoData = "Datos de configuración";
            }
            else if (OptionBtn1.Selected)
            {
                TipoData = "Master Data";
            }
            else if (OptionBtn2.Selected)
            {
                TipoData = "Datos de configuración";
            }

            string Query = string.Format(AddonMessageInfo.Query_List_Table_DTW, TipoData);

            SAPbouiCOM.DataTable oDT = oForm.GetDataTable("DT_DTW");
            Dow(Query, oDT);
            Grid1.CollapseLevel = 2;
            // Grid1.AssignLineNro();
            FormatoGridOBJ_DTW();
            oForm.Freeze(false);
        }
        public void FormatoGridOBJ_DTW()
        {
            Grid1.Columns.Item(0).TitleObject.Caption = labelsForms.LabelsFormsLabelTituloTablaDTW(Sb1Globals.Idioma); ;
            Grid1.Columns.Item(1).TitleObject.Caption = "";
            Grid1.Columns.Item(2).TitleObject.Caption = "";
            Grid1.Columns.Item(3).Visible = false;
            Grid1.AutoResizeColumns();
        }

        private void Grid1_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.Row >= 0)
                {
                    if (pVal.Row > 1)
                    {
                        Grid1.Rows.SelectedRows.Add(pVal.Row);
                        EditText0.Value = Grid1.DataTable.GetValue("Level_1_2_3", Grid1.GetDataTableRowIndex(pVal.Row)).ToString();
                        Obj_DTW = Grid1.DataTable.GetValue("ID_OBJ", Grid1.GetDataTableRowIndex(pVal.Row)).ToString();
                    }

                }
                if (pVal.ColUID == "Level_1_2_3")
                {
                    EditText0.Value = Grid1.DataTable.GetValue("Level_1_2_3", Grid1.GetDataTableRowIndex(pVal.Row)).ToString();
                    Obj_DTW = Grid1.DataTable.GetValue("ID_OBJ", Grid1.GetDataTableRowIndex(pVal.Row)).ToString();
                }

            }
            catch (Exception)
            {
                
            }
           
        }

        private SAPbouiCOM.Button Button8;

        private void Button8_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
           // const string FName = @"D:\William\AddDTW\Vistony.DTW.BO\UDO\PROMOCION\"; 
        }

        private void Button1_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            if (oForm.PaneLevel==4)
            {
                if (EditText0.Value=="")
                {
                    BubbleEvent = false;
                    Sb1Messages.ShowError("Debe Seleccionar al menos 1 Objecto");
                }
                else
                {

                    BubbleEvent = true;
                }
            }
            else
            {
                BubbleEvent = true;
            }
           
        }

        private void OptionBtn0_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            
        }

        private StaticText StaticText13;
        private EditText EditText0;
        private int SelectMatrix = 0;
        private void Matrix0_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {

            try
            {
                if (pVal.Row >= 0)
                {
                    if (pVal.Row >= 1)
                    {
                        Matrix0.SelectRow(pVal.Row,true,false);
                        SelectMatrix = pVal.Row;
                        // EditText0.Value = Grid1.DataTable.GetValue("ID_OBJ", Grid1.GetDataTableRowIndex(pVal.Row)).ToString();
                    }

                }

            }
            catch (Exception)
            {

            }

        }

        private Grid Grid2;
        private Grid Grid4;
        private Grid Grid5;
    }
}
