using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Funciones;
using eBilling;
using BOElectronicReception;
using SAPbobsCOM;
using SAPbouiCOM;
using System.IO;

namespace BOCore
{
    public class Core
    {
        #region Instanciacion Dll's

        Funciones.Comunes DllFunciones = new Funciones.Comunes();
        BOElectronicReception.ElectronicReception DllElectronicReception = new ElectronicReception();

        #endregion

        public void LoadParametersFormGestorAddOn(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Form _oFormGA)
        {
            try
            {

                #region Creacion Variables y Objetos

                string sGridAddInAvailable = null;
                string sGridAddInActive = null;
                string sLocalizacionActiva = null;
                string sSearchTable = null;                

                SAPbouiCOM.DataTable oDtAddInAvailable;
                SAPbouiCOM.DataTable oDtAddInActive;
                SAPbouiCOM.Grid oGridAddInAvailable;
                SAPbouiCOM.Grid oGridAddInActive;
                SAPbouiCOM.PictureBox oLogoBO;
                SAPbouiCOM.Folder oFolder1;
                SAPbouiCOM.StaticText olblVersion;
                SAPbouiCOM.ComboBox ocboLocalizacion;
                SAPbouiCOM.CheckBox ChkFEE;
                SAPbouiCOM.CheckBox ChkFER;

                SAPbobsCOM.Recordset oLisencedActiveAddOns = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                #endregion

                #region Instanciacion de variables y objetos

                oLogoBO = (SAPbouiCOM.PictureBox)(_oFormGA.Items.Item("oLogo").Specific);
                oDtAddInAvailable = _oFormGA.DataSources.DataTables.Add("oDtGridAD");
                oDtAddInActive = _oFormGA.DataSources.DataTables.Add("oDtGridAA");

                oGridAddInAvailable = (SAPbouiCOM.Grid)(_oFormGA.Items.Item("GridDispo").Specific);
                oGridAddInActive = (SAPbouiCOM.Grid)(_oFormGA.Items.Item("GridActi").Specific);

                oFolder1 = (SAPbouiCOM.Folder)_oFormGA.Items.Item("Folder1").Specific;

                olblVersion = (SAPbouiCOM.StaticText)(_oFormGA.Items.Item("lblVersion").Specific);

                ocboLocalizacion = (SAPbouiCOM.ComboBox)(_oFormGA.Items.Item("txtLoca").Specific);

                ChkFEE = (SAPbouiCOM.CheckBox)(_oFormGA.Items.Item("ChkFEE").Specific);
                ChkFER = (SAPbouiCOM.CheckBox)(_oFormGA.Items.Item("ChkFER").Specific);

                #endregion

                #region Consulta de AddIns Disponibles y cargue al formulario 

                sGridAddInAvailable = DllFunciones.GetStringXMLDocument(_oCompany, "Core", "ValidacionAddOnBO", "GridAddInAvailable");

                oDtAddInAvailable.ExecuteQuery(sGridAddInAvailable);

                oGridAddInAvailable.DataTable = oDtAddInAvailable;

                oGridAddInAvailable.AutoResizeColumns();

                #endregion

                #region Consulta de AddIns Activos y cargue al formulario 

                sGridAddInActive = DllFunciones.GetStringXMLDocument(_oCompany, "Core", "ValidacionAddOnBO", "GridAddInActive");

                oDtAddInActive.ExecuteQuery(sGridAddInActive);

                oGridAddInActive.DataTable = oDtAddInActive;

                oGridAddInActive.AutoResizeColumns();

                #endregion                

                #region Consulta Localizacion Activa

                SAPbobsCOM.Recordset oRsSearchTable = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oRsLocalizacion = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);                

                sSearchTable = DllFunciones.GetStringXMLDocument(_oCompany, "Core", "ValidacionAddOnBO", "SearchTables");
                sSearchTable = sSearchTable.Replace("%sTable%", "@BOEBILLINGP");                

                oRsSearchTable.DoQuery(sSearchTable);                

                if (oRsSearchTable.RecordCount > 0 )
                {
                    sLocalizacionActiva = DllFunciones.GetStringXMLDocument(_oCompany, "Core", "ValidacionAddOnBO", "LocalizacionActive");

                    oRsLocalizacion.DoQuery(sLocalizacionActiva);

                    if (oRsLocalizacion.RecordCount > 0)
                    {
                        ocboLocalizacion.Select(Convert.ToString(oRsLocalizacion.Fields.Item("Localizacion").Value.ToString()), BoSearchKey.psk_ByValue);
                    }
                    else
                    {

                    }
                }
                else
                {

                }

                #endregion


                #region Habilita Checkbox Addins

                ChkFEE.Item.Enabled = false;
                ChkFER.Item.Enabled = false;

                string LisencedActiveAddOns = DllFunciones.GetStringXMLDocument(_oCompany, "Core", "ValidacionAddOnBO", "LisencedActiveAddOns");

                oLisencedActiveAddOns.DoQuery(LisencedActiveAddOns);

                if (oLisencedActiveAddOns.RecordCount > 0)
                {
                    oLisencedActiveAddOns.MoveFirst();

                    do
                    {
                        string ConsultaAddin = null;

                        ConsultaAddin = Convert.ToString(oLisencedActiveAddOns.Fields.Item("AddIn").Value.ToString());

                        if (ConsultaAddin == "AddIneBillingBO")
                        {
                            ChkFEE.Item.Enabled = true;
                        }
                        else if (ConsultaAddin == "AddInElectronicReception")
                        {
                            ChkFER.Item.Enabled = true;
                        }

                        oLisencedActiveAddOns.MoveNext();

                    } while (oLisencedActiveAddOns.EoF == false);

                }

                #endregion

                #region Asignacion Logo

                oLogoBO.Picture = (Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Core\\Images\\LogoBO20x20.bmp");

                #endregion

                #region Asignacion Label

                olblVersion.Caption = "AddOn BOne ";
                olblVersion.Item.TextStyle = 1;

                #endregion

                _oFormGA.Refresh();

                _oFormGA.Visible = true;

                oFolder1.Select();

                

                



            }
            catch (Exception e)
            {

                DllFunciones.sendErrorMessage(_sboapp, e);
            }





        }

        public void OpenFormGestorAddOn(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _oCompany)
        {
            try
            {
                Form frm = null;
                bool ExistForm = false;

                for (int i = 0; i < _sboapp.Forms.Count; i++)
                {
                    if (_sboapp.Forms.Item(i).UniqueID == "BO_Gestion_AddOn")
                    {
                        frm = _sboapp.Forms.Item("BO_Gestion_AddOn");
                        ExistForm = true;
                    }
                }

                //si el formulario ya estaba creado le hace focus

                if (ExistForm)
                {
                    frm.Select();
                }
                else
                {
                    string ArchivoSRF = "Gestion_AddOn.srf";
                    DllFunciones.LoadFromXML(_sboapp, "Core", ref ArchivoSRF);

                    SAPbouiCOM.Form oFormGA;
                    oFormGA = _sboapp.Forms.Item("BO_Gestion_AddOn");

                    LoadParametersFormGestorAddOn(_sboapp, _oCompany, oFormGA);
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        public void TablasyCamposBaseBO(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company _company, string _sNameDB)
        {
            try
            {
                #region Variables y Objetos 

                string sQuerieValidacion = null;
                string sQuerieValidacionCopia = null;
                string sVersionAddInPresupuesto;
                string sVersionAddInIntercompany;
                string sVersionAddIneBillingBO;
                string sVersionAddInElectronicReception;
                string sVersionAddInProductionBO;
                string sVersionInstalador;

                #endregion

                #region Instanciacion Dlls
                                
                eBilling.eBillingBO DlleBilling = new eBilling.eBillingBO(sboapp, _company);                                
                BOElectronicReception.ElectronicReception DllElectronicReception = new ElectronicReception();                

                #endregion

                #region Version Dll's

                sVersionAddInPresupuesto = "2.0.0.0";
                sVersionAddInIntercompany = "2.0.0.0";                
                sVersionAddInProductionBO = "2.0.0.0";
                sVersionAddIneBillingBO = DlleBilling.VersionDll();
                sVersionAddInElectronicReception = DllElectronicReception.VersionDll();

                #endregion

                #region Creacion de tablas

                DllFunciones.crearTabla(_company, sboapp, "BOAdminAddOn", "Admin AddOn BOne", BoUTBTableType.bott_NoObject);

                #endregion

                #region Creacion de campos

                DllFunciones.CreaCamposUsr(_company, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 20, "", BoYesNoEnum.tNO, null, "BOAdminAddOn", "Version", "Version");
                DllFunciones.CreaCamposUsr(_company, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100, "", BoYesNoEnum.tNO, null, "BOAdminAddOn", "Licencia", "Licencia");

                string[] ValidValuesFields = { "A", "Activo", "I", "Inactivo", "S", "Sin Licencia" };
                DllFunciones.CreaCamposUsr(_company, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1, "", BoYesNoEnum.tNO, ValidValuesFields, "BOAdminAddOn", "Status", "Status");

                #endregion

                #region Insercion o actualizacion AddIns

                sQuerieValidacion = DllFunciones.GetStringXMLDocument(_company, "Core", "ValidacionAddOnBO", "ValidationExistingAddins");

                #region AddIn AddInPresupuesto

                SAPbobsCOM.Recordset oValidacionAddInPresupuesto = ((SAPbobsCOM.Recordset)_company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));

                sQuerieValidacionCopia = sQuerieValidacion.Replace("%AddIn%", "AddInPresupuesto");

                oValidacionAddInPresupuesto.DoQuery(sQuerieValidacionCopia);

                if (oValidacionAddInPresupuesto.RecordCount > 0)
                {
                    int sVersionAddInIgual;

                    sVersionAddInIgual = string.Compare(Convert.ToString(oValidacionAddInPresupuesto.Fields.Item("Version").Value.ToString()), sVersionAddInPresupuesto);

                    if (sVersionAddInIgual != 0)
                    {
                        DllFunciones.UpdateAddIn(sboapp, _company, "AddInPresupuesto", sVersionAddInPresupuesto);
                    }
                }
                else
                {
                    DllFunciones.InsertAddIn(sboapp, _company, "AddInPresupuesto", "AddInPresupuesto", sVersionAddInIntercompany, _sNameDB);
                }

                DllFunciones.liberarObjetos(oValidacionAddInPresupuesto);

                #endregion

                #region AddIn AddInIntercompany

                SAPbobsCOM.Recordset oValidacionAddInIntercompany = ((SAPbobsCOM.Recordset)_company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));

                sQuerieValidacionCopia = sQuerieValidacion.Replace("%AddIn%", "AddInIntercompany");

                oValidacionAddInIntercompany.DoQuery(sQuerieValidacionCopia);

                if (oValidacionAddInIntercompany.RecordCount > 0)
                {
                    int sVersionAddInIgual;

                    sVersionAddInIgual = string.Compare(Convert.ToString(oValidacionAddInIntercompany.Fields.Item("Version").Value.ToString()), sVersionAddInIntercompany);

                    if (sVersionAddInIgual != 0)
                    {
                        DllFunciones.UpdateAddIn(sboapp, _company, "AddInIntercompany", sVersionAddInIntercompany);
                    }
                }
                else
                {
                    DllFunciones.InsertAddIn(sboapp, _company, "AddInIntercompany", "AddInIntercompany", sVersionAddInIntercompany, _sNameDB);
                }

                DllFunciones.liberarObjetos(oValidacionAddInIntercompany);

                #endregion

                #region AddIn Facturacion Electronica

                SAPbobsCOM.Recordset oValidacionAddIneBillingBO = ((SAPbobsCOM.Recordset)_company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));

                sQuerieValidacionCopia = sQuerieValidacion.Replace("%AddIn%", "AddIneBillingBO");

                oValidacionAddIneBillingBO.DoQuery(sQuerieValidacionCopia);

                if (oValidacionAddIneBillingBO.RecordCount > 0)
                {
                    int sVersionAddInIgual;

                    sVersionAddInIgual = string.Compare(Convert.ToString(oValidacionAddIneBillingBO.Fields.Item("Version").Value.ToString()), sVersionAddIneBillingBO);

                    if (sVersionAddInIgual != 0)
                    {
                        DllFunciones.UpdateAddIn(sboapp, _company, "AddIneBillingBO", sVersionAddIneBillingBO);
                    }
                }
                else
                {
                    DllFunciones.InsertAddIn(sboapp, _company, "AddIneBillingBO", "AddIneBillingBO", sVersionAddInIntercompany, _sNameDB);
                }

                DllFunciones.liberarObjetos(oValidacionAddIneBillingBO);

                #endregion

                #region AddIn AddIneProductionBO

                SAPbobsCOM.Recordset oValidacionAddIneProductionBO = ((SAPbobsCOM.Recordset)_company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));

                sQuerieValidacionCopia = sQuerieValidacion.Replace("%AddIn%", "AddInProduccion");

                oValidacionAddIneProductionBO.DoQuery(sQuerieValidacionCopia);

                if (oValidacionAddIneProductionBO.RecordCount > 0)
                {
                    int sVersionAddInIgual;

                    sVersionAddInIgual = string.Compare(Convert.ToString(oValidacionAddIneProductionBO.Fields.Item("Version").Value.ToString()), sVersionAddInProductionBO);

                    if (sVersionAddInIgual != 0)
                    {
                        DllFunciones.UpdateAddIn(sboapp, _company, "AddInProduccion", sVersionAddInProductionBO);
                    }
                }
                else
                {
                    DllFunciones.InsertAddIn(sboapp, _company, "AddInProduccion", "Produccion Avanzada", sVersionAddInProductionBO, _sNameDB);
                }

                DllFunciones.liberarObjetos(oValidacionAddIneProductionBO);

                #endregion

                #region AddIn AddInElectronicReception

                SAPbobsCOM.Recordset oValidacionAddInElectronicReception = ((SAPbobsCOM.Recordset)_company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));

                sQuerieValidacionCopia = sQuerieValidacion.Replace("%AddIn%", "AddInElectronicReception");

                oValidacionAddInElectronicReception.DoQuery(sQuerieValidacionCopia);

                if (oValidacionAddInElectronicReception.RecordCount > 0)
                {
                    int sVersionAddInIgual;

                    sVersionAddInIgual = string.Compare(Convert.ToString(oValidacionAddInElectronicReception.Fields.Item("Version").Value.ToString()), sVersionAddInElectronicReception);

                    if (sVersionAddInIgual != 0)
                    {
                        DllFunciones.UpdateAddIn(sboapp, _company, "AddInElectronicReception", sVersionAddIneBillingBO);
                    }
                }
                else
                {
                    DllFunciones.InsertAddIn(sboapp, _company, "AddInElectronicReception", "Recepcion Electronica", sVersionAddInElectronicReception, _sNameDB);
                }

                DllFunciones.liberarObjetos(oValidacionAddInElectronicReception);

                #endregion

                #endregion

            }
            catch (Exception e)
            {

                DllFunciones.sendErrorMessage(sboapp, e);
            }
        }

        public void ValidacionEstructuraAddInsActivos(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company _company, string _sNameDB)
        {
            #region Variables y objetos 

            SAPbouiCOM.Form oBO_Gestion_AddOn = sboapp.Forms.Item("BO_Gestion_AddOn");
            SAPbouiCOM.CheckBox oChkFEE = (SAPbouiCOM.CheckBox)oBO_Gestion_AddOn.Items.Item("ChkFEE").Specific;
            SAPbouiCOM.CheckBox oChkFER = (SAPbouiCOM.CheckBox)oBO_Gestion_AddOn.Items.Item("ChkFER").Specific;

            SAPbobsCOM.Recordset oConsultaAddIns = (SAPbobsCOM.Recordset)_company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string sLocalizacion;
            string sConsultaAddins;
            int ContadorAddInsActive = 0;
            
            #endregion

            #region Instanciacion Dll's

            eBilling.eBillingBO DllFacturaionElectronica= new eBilling.eBillingBO(sboapp, _company);

            #endregion

            #region ConsultaAddIns

            sConsultaAddins = DllFunciones.GetStringXMLDocument(_company, "Core", "ValidacionAddOnBO", "AddInsActive");

            oConsultaAddIns.DoQuery(sConsultaAddins);

            if (oConsultaAddIns.RecordCount > 0)
            {
                oConsultaAddIns.MoveFirst();

                do
                {
                    sConsultaAddins = null;

                    sConsultaAddins = Convert.ToString(oConsultaAddIns.Fields.Item("Code").Value.ToString());

                    if (sConsultaAddins == "AddIneBillingBO")
                    {
                        oChkFEE.Item.Visible = true;
                    }
                    else if (sConsultaAddins == "AddInElectronicReception")
                    {
                        oChkFER.Item.Visible = true;
                    }                    

                    oConsultaAddIns.MoveNext();

                } while (oConsultaAddIns.EoF == false);
            }


            DllFunciones.liberarObjetos(oConsultaAddIns);

            #endregion

            #region Consulta Localizacion 

            sLocalizacion = ((SAPbouiCOM.ComboBox)(oBO_Gestion_AddOn.Items.Item("txtLoca").Specific)).Value.ToString();

            #endregion

            #region Validacion de tablas y campos 

            if (string.IsNullOrEmpty(sLocalizacion))
            {
                DllFunciones.sendMessageBox(sboapp, "No ha seleccionado el tipo de localización, por favor seleccionar para poder continuar");
            }
            else
            {
                

                

                if (oChkFEE.Checked)
                {
                    ContadorAddInsActive++;
                }

                if (oChkFER.Checked)
                {
                    ContadorAddInsActive++;
                }

                if (ContadorAddInsActive == 0)
                {
                    DllFunciones.sendMessageBox(sboapp, "Por favor seleccionar como minimo 1 modulo para validar estructura ");
                }
                else
                {
                    int sProcesar = DllFunciones.sendMessageBoxY_N(sboapp, "Se validaran las tablas y campos de todos los Add-Ins seleccionados , ¿ Desea continuar ?");

                    if (sProcesar == 1)
                    {
                        if (oChkFEE.Checked)
                        {
                            DllFacturaionElectronica.CreacionTablasyCamposeBillingBO(sboapp, _company, Convert.ToString(_company.DbServerType), sLocalizacion);
                        }

                        if (oChkFER.Checked)
                        {
                            DllElectronicReception.CreacionTablasyCamposeBillingBO(sboapp, _company);
                        }

                        DllFunciones.sendMessageBox(sboapp, "Se crearon las tablas y campos correctamente, por favor salir y volver a ingresar");
                    }
                }                
            }

            #endregion

        }
        
    }
}
