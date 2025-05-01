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
using System.Diagnostics;
using System.Drawing;
using System.Net.Http;
using Newtonsoft.Json;
using Presupuesto;
using Intercompany;
using BOProduccion;
using System.Security.Cryptography;

namespace BOCore
{
    public class Core
    {
        private static readonly HttpClient client = new HttpClient();

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
                SAPbouiCOM.StaticText olblEstado;
                SAPbouiCOM.StaticText olblUpdFea;
                SAPbouiCOM.StaticText olblDevelp;
                SAPbouiCOM.StaticText olblHelp;
                SAPbouiCOM.ComboBox ocboLocalizacion;
                SAPbouiCOM.CheckBox ChkFEE;
                SAPbouiCOM.CheckBox ChkFER;
                SAPbouiCOM.EditText otxtServer;

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
                olblDevelp = (SAPbouiCOM.StaticText)(_oFormGA.Items.Item("lblDevelp").Specific);
                olblEstado = (SAPbouiCOM.StaticText)(_oFormGA.Items.Item("lblEstado").Specific);
                olblUpdFea = (SAPbouiCOM.StaticText)(_oFormGA.Items.Item("lblUpdFea").Specific);
                olblHelp = (SAPbouiCOM.StaticText)(_oFormGA.Items.Item("lblHelp").Specific);

                ocboLocalizacion = (SAPbouiCOM.ComboBox)(_oFormGA.Items.Item("txtLoca").Specific);

                ChkFEE = (SAPbouiCOM.CheckBox)(_oFormGA.Items.Item("ChkFEE").Specific);
                ChkFER = (SAPbouiCOM.CheckBox)(_oFormGA.Items.Item("ChkFER").Specific);

                otxtServer = (SAPbouiCOM.EditText)(_oFormGA.Items.Item("txtServer").Specific);

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

                #region Asignacion Valores Label

                olblVersion.Caption = VersionAddOn();
                olblVersion.Item.TextStyle = 1;
                                
                olblDevelp.Item.TextStyle = 1;

                #region Obtiene si el sistema esta registrado

                SAPbobsCOM.Recordset oRsNITCompany = ((SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));

                string sNITCompany = DllFunciones.GetStringXMLDocument(_oCompany, "Core", "ValidacionAddOnBO", "NITCompany");

                oRsNITCompany.DoQuery(sNITCompany);

                if (oRsNITCompany.RecordCount > 0)
                {
                    sNITCompany = Convert.ToString(oRsNITCompany.Fields.Item("TaxIdNum").Value.ToString());
                    var responseRequestAPIHX = GetContractStatusAsync(sNITCompany).GetAwaiter().GetResult();

                    if (responseRequestAPIHX == "Error: NotFound - Not Found")
                    {
                        olblEstado.Caption = MessageSystem("CO-00004", _sboapp);
                        olblEstado.Item.ForeColor = ColorTranslator.ToOle(Color.Red);
                    }
                    else if (responseRequestAPIHX.Contains("Inactivo"))
                    {
                        olblEstado.Caption = responseRequestAPIHX.ToString();
                        olblEstado.Item.ForeColor = ColorTranslator.ToOle(Color.Red);
                    }
                    else
                    {
                        olblEstado.Caption = responseRequestAPIHX.ToString();
                        olblEstado.Item.ForeColor = ColorTranslator.ToOle(Color.Green);
                    }
                    
                }

                DllFunciones.liberarObjetos(oRsNITCompany);

                #endregion

                olblUpdFea.Item.ForeColor = ColorTranslator.ToOle(Color.Blue);
                olblUpdFea.Item.TextStyle = 4;

                olblHelp.Item.ForeColor = ColorTranslator.ToOle(Color.Blue);
                olblHelp.Item.TextStyle = 4;

                #endregion

                #region Consulta Servidor

                otxtServer.Value = _sboapp.Company.ServerName.ToString();

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

        public void CloseFormGestorAddOn(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _oCompany)
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
                    frm.Close();
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
                string sVersionAddInPresupuesto;
                string sVersionAddInIntercompany;
                string sVersionAddIneBillingBO;
                string sVersionAddInElectronicReception;
                string sVersionAddInProductionBO;
                string sVersionBillingConsolidator;

                #endregion

                #region Instanciacion Dlls

                eBilling.eBillingBO DlleBilling = new eBilling.eBillingBO(sboapp, _company);
                BOElectronicReception.ElectronicReception DllElectronicReception = new ElectronicReception();
                Presupuesto.Core DllPresupuesto = new Presupuesto.Core(sboapp, _company);
                Intercompany.Intercompany DllIntercompany = new Intercompany.Intercompany(sboapp, _company);
                BOProduccion.Production DllProduction = new Production();
                BOBillingConsolidator.BillingConsolidator DllBillingConsolidator = new BOBillingConsolidator.BillingConsolidator();

                #endregion

                #region Version Dll's

                sVersionAddInPresupuesto = DllPresupuesto.VersionDll();
                sVersionAddInIntercompany = DllIntercompany.VersionDll();
                sVersionAddInProductionBO = DllProduction.VersionDll();
                sVersionAddIneBillingBO = DlleBilling.VersionDll();
                sVersionAddInElectronicReception = DllElectronicReception.VersionDll();
                sVersionBillingConsolidator = DllBillingConsolidator.VersionAddOn();

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

                #region AddIn Presupuesto

                ValidationAddIn(_company, sboapp, "AddInPresupuesto", sVersionAddInPresupuesto, _sNameDB);

                #endregion

                #region AddIn Intercompany

                ValidationAddIn(_company, sboapp, "AddInIntercompany", sVersionAddInIntercompany, _sNameDB);

                #endregion

                #region AddIn Facturacion Electronica

                ValidationAddIn(_company, sboapp, "AddIneBillingBO", sVersionAddIneBillingBO, _sNameDB);

                #endregion

                #region AddIn Produccion Avanzada

                ValidationAddIn(_company, sboapp, "AddInProduccion", sVersionAddInProductionBO, _sNameDB);

                #endregion

                #region AddIn Recepcion Electronica

                ValidationAddIn(_company, sboapp, "AddInElectronicReception", sVersionAddInElectronicReception, _sNameDB);

                #endregion

                #region AddIn Consolidador de Facturación

                ValidationAddIn(_company, sboapp, "AddInBillingConsolidator", sVersionBillingConsolidator, _sNameDB);

                #endregion

                #endregion
                
            }
            catch (Exception e)
            {

                DllFunciones.sendErrorMessage(sboapp, e);
            }
        }

        public void ValidationAddIn(SAPbobsCOM.Company _company, SAPbouiCOM.Application sboapp, string sNombreAddIn, string sVersionAddIn, string _sNameDB)
        {
            string sQuerieValidacion = DllFunciones.GetStringXMLDocument(_company, "Core", "ValidacionAddOnBO", "ValidationExistingAddins");

            SAPbobsCOM.Recordset oValidacionAddIn = ((SAPbobsCOM.Recordset)_company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));

            string sQuerieValidacionCopia = sQuerieValidacion.Replace("%AddIn%", sNombreAddIn);

            oValidacionAddIn.DoQuery(sQuerieValidacionCopia);

            if (oValidacionAddIn.RecordCount > 0)
            {
                int sVersionAddInIgual;

                sVersionAddInIgual = string.Compare(Convert.ToString(oValidacionAddIn.Fields.Item("Version").Value.ToString()), sVersionAddIn);

                if (sVersionAddInIgual != 0)
                {
                    DllFunciones.UpdateAddIn(sboapp, _company, sNombreAddIn, sVersionAddIn);
                }
            }
            else
            {
                DllFunciones.InsertAddIn(sboapp, _company, sNombreAddIn, sNombreAddIn, sVersionAddIn, _sNameDB);
            }

            DllFunciones.liberarObjetos(oValidacionAddIn);

        }

        public void ValidacionEstructuraAddInsActivos(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company _company, string _sNameDB)
        {
            #region Variables y objetos 

            SAPbouiCOM.Form oBO_Gestion_AddOn = sboapp.Forms.Item("BO_Gestion_AddOn");
            SAPbouiCOM.CheckBox oChkFEE = (SAPbouiCOM.CheckBox)oBO_Gestion_AddOn.Items.Item("ChkFEE").Specific;
            SAPbouiCOM.CheckBox oChkFER = (SAPbouiCOM.CheckBox)oBO_Gestion_AddOn.Items.Item("ChkFER").Specific;

            bool bValidarCamposFacturacionElectronicaEmision = false;
            bool bValidarCamposRepcecionElectronica = false;

            string sLocalizacion;            
            int ContadorAddInsActive = 0;
            
            #endregion

            #region Instanciacion Dll's

            eBilling.eBillingBO DllFacturacionElectronica= new eBilling.eBillingBO(sboapp, _company);

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
                    bValidarCamposFacturacionElectronicaEmision = true;
                }

                if (oChkFER.Checked)
                {
                    ContadorAddInsActive++;
                    bValidarCamposRepcecionElectronica = true;
                }

                if (ContadorAddInsActive == 0)
                {
                    DllFunciones.sendMessageBox(sboapp, "Por favor seleccionar como minimo 1 AddIn para validar estructura ");
                }
                else
                {
                    int sProcesar = DllFunciones.sendMessageBoxY_N(sboapp, "Se validaran las tablas y campos de todos los Add-Ins seleccionados , ¿ Desea continuar ?");

                    if (sProcesar == 1)
                    {
                        if (bValidarCamposFacturacionElectronicaEmision)
                        {
                            DllFacturacionElectronica.CreacionTablasyCamposeBillingBO(sboapp, _company, Convert.ToString(_company.DbServerType), sLocalizacion);
                        }

                        if (bValidarCamposRepcecionElectronica)
                        {
                            DllElectronicReception.CreacionTablasyCamposeBillingBO(sboapp, _company);
                        }

                        DllFunciones.sendMessageBox(sboapp, "Se crearon las tablas y campos correctamente, por favor salir y volver a ingresar");
                    }
                }                
            }

            #endregion

        }

        public string MessageSystem(string idMessage, Application sboAppCore)
        {
            if (idMessage == "00001")
            {
                if (sboAppCore.Language == BoLanguages.ln_Spanish_La)
                {
                    return idMessage + " - Cargando AddOn BOne ,  espere por favor....";
                }
                else
                {
                    return idMessage + " - Loading AddOn BOne, please wait....";
                }
            }
            else if (idMessage == "00002")
            {
                if (sboAppCore.Language == BoLanguages.ln_Spanish_La)
                {
                    return idMessage + " - AddOn BOne cargado correctamente";
                }
                else
                {
                    return idMessage + " - AddOn BOne loaded successfully";
                }
            }
            else if (idMessage == "00003")
            {
                if (sboAppCore.Language == BoLanguages.ln_Spanish_La)
                {
                    return idMessage + " - El archivo no existe el archivo \"Ultimas mejoras y actualizaciones\" ";
                }
                else
                {
                    return idMessage + " - The file does not exist the file \"Latest improvements and updates\"";
                }
            }
            else if (idMessage == "CO-00004")
            {
                if (sboAppCore.Language == BoLanguages.ln_Spanish_La)
                {
                    return idMessage + " - Su sistema NO se encuentra registrado";
                }
                else
                {
                    return idMessage + " - Your system is NOT registered";
                }
            }
            else
            {
                return "";
            }
            
        }

        public void OpenFileUpdatesAndImprovements(SAPbouiCOM.Application sboAppCore, SAPbobsCOM.Company oCompanyCore)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            string filePath = AppDomain.CurrentDomain.BaseDirectory + "Core\\Docs\\Updates_and_improvements\\Index.html";

            if (File.Exists(filePath))
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true,
                    WindowStyle = ProcessWindowStyle.Normal
                });
            }
            else
            {
                DllFunciones.sendMessageBox(sboAppCore, MessageSystem("00003", sboAppCore));                
            }
        }

        public void OpenURLHelpDesk(SAPbouiCOM.Application sboAppCore, SAPbobsCOM.Company oCompanyCore)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            string filePath = AppDomain.CurrentDomain.BaseDirectory + "Core\\Docs\\Updates_and_improvements\\Index.html";

            if (File.Exists(filePath))
            {
                Process.Start(new ProcessStartInfo("https://hexagramaconsulting.atlassian.net/servicedesk/customer/portal/1")
                {                    
                    UseShellExecute = true,
                    WindowStyle = ProcessWindowStyle.Normal
                });
            }
            else
            {
                DllFunciones.sendMessageBox(sboAppCore, MessageSystem("00003", sboAppCore));
            }
        }

        public string Encriptar(string texto, string Llave)
        {
            try
            {
                byte[] keyArray;

                byte[] Arreglo_a_Cifrar = UTF8Encoding.UTF8.GetBytes(texto);

                //Se utilizan las clases de encriptación MD5

                MD5CryptoServiceProvider hashmd5 = new MD5CryptoServiceProvider();

                keyArray = hashmd5.ComputeHash(UTF8Encoding.UTF8.GetBytes(Llave));

                hashmd5.Clear();

                //Algoritmo TripleDES
                TripleDESCryptoServiceProvider tdes = new TripleDESCryptoServiceProvider();

                tdes.Key = keyArray;
                tdes.Mode = CipherMode.ECB;
                tdes.Padding = PaddingMode.PKCS7;

                ICryptoTransform cTransform = tdes.CreateEncryptor();

                byte[] ArrayResultado = cTransform.TransformFinalBlock(Arreglo_a_Cifrar, 0, Arreglo_a_Cifrar.Length);

                tdes.Clear();

                //se regresa el resultado en forma de una cadena
                texto = Convert.ToBase64String(ArrayResultado, 0, ArrayResultado.Length);

            }
            catch (Exception)
            {

            }
            return texto;
        }

        private async Task<string> GetContractStatusAsync(string NITEmpresa)
        {
            try
            {
                // URL de tu API
                string apiUrl = $"https://hexagramaconsulting.com/wp-json/hexagrama/v1/contracts/{NITEmpresa}";

                // Realizar la solicitud GET
                HttpResponseMessage response = await client.GetAsync(apiUrl);

                // Verificar si la respuesta es exitosa
                if (response.IsSuccessStatusCode)
                {
                    string responseData = await response.Content.ReadAsStringAsync();

                    var requestObject = JsonConvert.DeserializeObject<Contract>(responseData);

                    return $"Contrato No. : {requestObject.ContractId}, Estado: {requestObject.Status}, Fecha Finalizacion: {requestObject.EndDate}" ;

                }
                else
                {
                    string responseData = $"Error: {response.StatusCode} - {response.ReasonPhrase}";

                    return responseData;
                }
            }
            catch (Exception ex)
            {
                return $"Error al consumir la API: {ex.Message}";                
            }
        }
        
        #region Clases JSON

        public class Contract
        {
            public string ContractId { get; set; }
            public string ClientName { get; set; }
            public string EndDate { get; set; }
            public string Status { get; set; }

        }

        #endregion

        private string VersionAddOn()
        {
            try
            {             
                FileVersionInfo myFileVersionInfo = FileVersionInfo.GetVersionInfo(AppDomain.CurrentDomain.BaseDirectory + "BasisOne.exe");

                return myFileVersionInfo.FileVersion.ToString();
            }
            catch (Exception)
            {

                throw;
            }

        }


    }
}
