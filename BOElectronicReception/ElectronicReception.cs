using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using SAPbouiCOM;
using System.Reflection;
using Funciones;
using BOElectronicReception.Recepcion21WS;
using System.ServiceModel;
using System.IO;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Net;
using System.Data;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Xml;
using System.Diagnostics;
using ExcelHX = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;


namespace BOElectronicReception
{
    public class ElectronicReception
    {
        private readonly HttpClient _httpClient;

        #region Instanciacion

        Funciones.Comunes DllFunciones = new Funciones.Comunes();

        #endregion

        #region Parametros globales TFHKA

        Recepcion21WS.ReceptorWSClient RecepcionWS;

        BasicHttpBinding port = new BasicHttpBinding();

        #endregion

        public ElectronicReception()
        {            
            _httpClient = new HttpClient();
        }
        
        public void OpenFormVisorDocumentosRecibidos(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _oCompany, string sMotor)
        {
            try
            {
                SAPbouiCOM.Form frm = null;
                bool ExistForm = false;

                for (int i = 0; i < _sboapp.Forms.Count; i++)
                {
                    if (_sboapp.Forms.Item(i).UniqueID == "BOTVDR")
                    {
                        frm = _sboapp.Forms.Item("BOTVDR");
                        ExistForm = true;
                    }
                }


                if (ExistForm)
                {
                    frm.Select();
                }
                else
                {
                    string ArchivoSRF = "VisorDocumentsReceptionElectronic.srf";
                    DllFunciones.LoadFromXML(_sboapp, "BOElectronicReception", ref ArchivoSRF);

                    frm = _sboapp.Forms.Item("BOTVDR");

                    LoadFormDocumentsReception(_sboapp, _oCompany, frm);

                }
            }
            catch (Exception)
            {

                throw;
            }

        }

        public void CreacionTablasyCamposeBillingBO(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {
                #region Creacion de tablas

                //1
                DllFunciones.ProgressBar(oCompany, sboapp, 30, 1, "Creando Tabla - Recepcion Electronica - Documentos recibidos PT, por favor espere...");
                DllFunciones.crearTabla(oCompany, sboapp, "BOTREDRPT", "BO Doc. Rec. Prov", SAPbobsCOM.BoUTBTableType.bott_NoObject);

                #endregion

                #region Creacion Campos Tabla recepcion electronica documentos recibidos por proveedor tecnologico

                DllFunciones.ProgressBar(oCompany, sboapp, 30, 1, "Creando Campo - Recep. Electro. - PT, por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 10, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTPT", "Proveedor Technologico");

                DllFunciones.ProgressBar(oCompany, sboapp, 30, 1, "Creando Campo - Recep. Electro. - Numero Documento PT, por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 10, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTNDPT", "Numero Documento PT");

                DllFunciones.ProgressBar(oCompany, sboapp, 30, 1, "Creando Campo - Recep. Electro. - CUFE, por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTCUFE", "CUFE");

                DllFunciones.ProgressBar(oCompany, sboapp, 30, 1, "Creando Campo - Recep. Electro. - Status DIAN Codigo, por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 10, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTSDC", "Status DIAN Codi.");
                
                DllFunciones.ProgressBar(oCompany, sboapp, 30, 1, "Creando Campo - Recep. Electro. - Status DIAN Descripc., por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 10, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTSDD", "Status DIAN Desc.");

                DllFunciones.ProgressBar(oCompany, sboapp, 30, 1, "Creando Campo - Recep. Electro. - Status DIAN Fecha., por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 25, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTSDF", "Status DIAN Fech.");

                DllFunciones.ProgressBar(oCompany, sboapp, 30, 1, "Creando Campo - Recep. Electro. - Fecha Emision, por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 25, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTFE", "Fecha Emision.");

                DllFunciones.ProgressBar(oCompany, sboapp, 30, 1, "Creando Campo - Recep. Electro. - Fecha Recepcion, por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 25, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTFR", "Fecha Recep.");

                DllFunciones.ProgressBar(oCompany, sboapp, 30, 1, "Creando Campo - Recep. Electro. - Hora Emision, por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 25, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTHE", "Hora Emision.");

                DllFunciones.ProgressBar(oCompany, sboapp, 30, 1, "Creando Campo - Recep. Electro. - Monto Total, por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Float, BoFldSubTypes.st_Price, 11, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTMT", "Monto Total");

                DllFunciones.ProgressBar(oCompany, sboapp, 30, 1, "Creando Campo - Recep. Electro. - Numero Factura, por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 50, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTNF", "Numero Factura");

                DllFunciones.ProgressBar(oCompany, sboapp, 30, 1, "Creando Campo - Recep. Electro. - Numero Identificacion, por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 20, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTNI", "(HX)Numero Identificacion");

                DllFunciones.ProgressBar(oCompany, sboapp, 30, 1, "Creando Campo - Recep. Electro. - Razon Social, por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTNRS", "(HX)Razon Social");

                DllFunciones.ProgressBar(oCompany, sboapp, 30, 1, "Creando Campo - Recep. Electro. - Tipo Documento, por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 2, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTTD", "(HX)Tipo Documento");

                DllFunciones.ProgressBar(oCompany, sboapp, 30, 1, "Creando Campo - Recep. Electro. - Tipo Emisor, por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 2, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTTE", "(HX)Tipo Emisor");

                DllFunciones.ProgressBar(oCompany, sboapp, 30, 1, "Creando Campo - Recep. Electro. - Tipo Identidad, por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 3, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTTI", "(HX)Tipo Identidad");

                DllFunciones.ProgressBar(oCompany, sboapp, 30, 1, "Creando Campo - Estado Documento DIAN - Tipo Identidad, por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 3, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTEDD", "(HX)Est. Doc. DIAN");

                DllFunciones.ProgressBar(oCompany, sboapp, 30, 1, "Creando Campo - Numero Documento Preeliminar , por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 3, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTNDP", "(HX)Num. Doc. Preel");

                DllFunciones.ProgressBar(oCompany, sboapp, 30, 1, "Creando Campo - Numero Documento Definitivo , por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 3, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTNDD", "(HX)Num. Doc. Def.");

                DllFunciones.ProgressBar(oCompany, sboapp, 30, 1, "Creando Campo - Estado Evento DIAN , por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 3, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTEED", "(HX)Est. Even. DIAN");

                DllFunciones.ProgressBar(oCompany, sboapp, 30, 1, "Creando Campo - Path Adjunto XML , por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Memo, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTPAXML", " (HX)Path. Adju. XML");

                DllFunciones.ProgressBar(oCompany, sboapp, 30, 1, "Creando Campo - Codigo Respuesta WS XML , por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 10, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTCWSXML", " (HX)Cod. Resp. WS XML");

                DllFunciones.ProgressBar(oCompany, sboapp, 30, 1, "Creando Campo - Path Adjunto PDF , por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Memo, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTPAPDF", " (HX)Path. Adju. PDF");

                DllFunciones.ProgressBar(oCompany, sboapp, 30, 1, "Creando Campo - Codigo Respuesta WS PDF , por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 10, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTCWSPDF", "(HX) Cod. Resp. WS PDF");

                DllFunciones.ProgressBar(oCompany, sboapp, 30, 1, "Creando Campo - Tipo de documento , por favor espere ...");
                string[] ValidValuesFields1 = { "13", "Cedula Ciudadania", "31", "NIT" };
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 10, "", BoYesNoEnum.tNO, ValidValuesFields1, "OHEM", "BOTTD", "(HX) Tipo documento");

                DllFunciones.ProgressBar(oCompany, sboapp, 30, 1, "Creando Campo - Numero de Orden de Compra, por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 30, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "HXPO", "(HX) Orden Compra");

                DllFunciones.ProgressBar(oCompany, sboapp, 30, 1, "Creando Campo - Numero de Orden de Compra, por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 30, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "HXPAYME", "(HX) Condicion Pago ");


                #endregion

                #region Creacion campos tabla parametros generales 

                DllFunciones.ProgressBar(oCompany, sboapp, 30, 1, "Creando Campo - Forma de Recepción, por favor espere...");
                string[] ValidValuesFormaRecepcion = { "A", "Documentos", "B", "Documentos y XML", "C", "Documentos, XML y PDF" };
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 2, "", BoYesNoEnum.tNO, ValidValuesFormaRecepcion, "@BOEBILLINGP", "HX_FRE", "(HX) Forma Recep. Elec.");

                #endregion

                #region Crea procedimientos almacenados

                DllFunciones.ProgressBar(oCompany, sboapp, 30, 1, "Creando Procedimientos Almacenados Por favor espere ...");

                SAPbobsCOM.Recordset oProcedures = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string sProcedure_Eliminar = null;
                string sProcedure_Crear = null;

                if (Convert.ToString(oCompany.DbServerType) == "dst_HANADB")
                {
                    #region Consulta si Existente el Procedure BOT_InsertDocuments

                    sProcedure_Eliminar = DllFunciones.GetStringXMLDocument(oCompany, "BOElectronicReception", "ElectronicReception", "SearchProcedure");
                    sProcedure_Eliminar = sProcedure_Eliminar.Replace("%sNameProcedure%", "BOT_InsertDocuments");

                    oProcedures.DoQuery(sProcedure_Eliminar);

                    #endregion

                    #region Crea el procedimiento almacenado BOT_InsertDocuments

                    if (oProcedures.RecordCount > 0)
                    {
                        #region Elimina el procedure 

                        sProcedure_Eliminar = null;
                        sProcedure_Eliminar = DllFunciones.GetStringXMLDocument(oCompany, "BOElectronicReception", "ElectronicReception", "Eliminar_BO_FacturaXML");
                        sProcedure_Eliminar = sProcedure_Eliminar.Replace("%sNameProcedure%", "BOT_InsertDocuments");

                        oProcedures.DoQuery(sProcedure_Eliminar);

                        #endregion                        
                    }

                    sProcedure_Crear = DllFunciones.GetStringXMLDocument(oCompany, "BOElectronicReception", "ElectronicReception", "BOT_InsertDocuments");                    

                    oProcedures.DoQuery(sProcedure_Crear);

                    #endregion

                    #region Consulta si Existente el Procedure BOT_SyncAttachment

                    sProcedure_Eliminar = DllFunciones.GetStringXMLDocument(oCompany, "BOElectronicReception", "ElectronicReception", "SearchProcedure");
                    sProcedure_Eliminar = sProcedure_Eliminar.Replace("%sNameProcedure%", "BOT_SyncAttachment");

                    oProcedures.DoQuery(sProcedure_Eliminar);

                    if (oProcedures.RecordCount > 0)
                    {
                        #region Elimina el procedure 

                        sProcedure_Eliminar = null;
                        sProcedure_Eliminar = DllFunciones.GetStringXMLDocument(oCompany, "BOElectronicReception", "ElectronicReception", "Eliminar_BO_FacturaXML");
                        sProcedure_Eliminar = sProcedure_Eliminar.Replace("%sNameProcedure%", "BOT_SyncAttachment");

                        oProcedures.DoQuery(sProcedure_Eliminar);

                        #endregion                        
                    }

                    #endregion

                    #region Crea el procedimiento almacenado BOT_SyncAttachment

                    sProcedure_Crear = DllFunciones.GetStringXMLDocument(oCompany, "BOElectronicReception", "ElectronicReception", "BOT_SyncAttachment");

                    oProcedures.DoQuery(sProcedure_Crear);

                    #endregion

                    #region Consulta si Existente el Procedure BOT_DIGVER

                    sProcedure_Eliminar = DllFunciones.GetStringXMLDocument(oCompany, "BOElectronicReception", "ElectronicReception", "SearchProcedure");
                    sProcedure_Eliminar = sProcedure_Eliminar.Replace("%sNameProcedure%", "BOTDIGVER");

                    oProcedures.DoQuery(sProcedure_Eliminar);

                    if (oProcedures.RecordCount > 0)
                    {
                        #region Elimina el procedure 

                        sProcedure_Eliminar = null;
                        sProcedure_Eliminar = DllFunciones.GetStringXMLDocument(oCompany, "BOElectronicReception", "ElectronicReception", "Eliminar_BO_FacturaXML");
                        sProcedure_Eliminar = sProcedure_Eliminar.Replace("%sNameProcedure%", "BOTDIGVER");

                        oProcedures.DoQuery(sProcedure_Eliminar);

                        #endregion                        
                    }

                    #endregion

                    #region Crea el procedimiento almacenado BOT_DIGVER

                    sProcedure_Crear = DllFunciones.GetStringXMLDocument(oCompany, "BOElectronicReception", "ElectronicReception", "BOT_DIGVER");

                    oProcedures.DoQuery(sProcedure_Crear);

                    #endregion

                }
                else
                {
                    #region Consulta si el procedure Existe BOT_InsertDocuments

                    sProcedure_Eliminar = DllFunciones.GetStringXMLDocument(oCompany, "BOElectronicReception", "ElectronicReception", "Eliminar_BO_FacturaXML");
                    sProcedure_Eliminar = sProcedure_Eliminar.Replace("%sNameProcedure%", "BO_FacturaXML");

                    oProcedures.DoQuery(sProcedure_Eliminar);

                    #endregion

                    #region Crea el procedimiento almacenado BOT_InsertDocuments

                    sProcedure_Crear = DllFunciones.GetStringXMLDocument(oCompany, "BOElectronicReception", "ElectronicReception", "BOT_InsertDocuments");

                    oProcedures.DoQuery(sProcedure_Crear);

                    #endregion

                    #region Consulta si el procedure Existe BOT_SyncAttachment

                    sProcedure_Eliminar = DllFunciones.GetStringXMLDocument(oCompany, "BOElectronicReception", "ElectronicReception", "Eliminar_BO_FacturaXML");
                    sProcedure_Eliminar = sProcedure_Eliminar.Replace("%sNameProcedure%", "BOT_SyncAttachment");

                    oProcedures.DoQuery(sProcedure_Eliminar);

                    #endregion

                    #region Crea el procedimiento almacenado BOT_SyncAttachment

                    sProcedure_Crear = DllFunciones.GetStringXMLDocument(oCompany, "BOElectronicReception", "ElectronicReception", "BOT_SyncAttachment");

                    oProcedures.DoQuery(sProcedure_Crear);

                    #endregion

                    #region Consulta si el procedure Existe BOT_DIGVER

                    sProcedure_Eliminar = DllFunciones.GetStringXMLDocument(oCompany, "BOElectronicReception", "ElectronicReception", "Eliminar_BO_FacturaXML");
                    sProcedure_Eliminar = sProcedure_Eliminar.Replace("%sNameProcedure%", "BOTDIGVER");

                    oProcedures.DoQuery(sProcedure_Eliminar);

                    #endregion

                    #region Crea el procedimiento almacenado BOT_DIGVER

                    sProcedure_Crear = DllFunciones.GetStringXMLDocument(oCompany, "BOElectronicReception", "ElectronicReception", "BOT_DIGVER");

                    oProcedures.DoQuery(sProcedure_Crear);

                    #endregion

                }



                #endregion

            }
            catch (Exception e)
            {
                DllFunciones.sendErrorMessage(sboapp, e);
            }

        }        

        public void LoadFormDocumentsReception(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oFormReception)
        {
            Funciones.Comunes DLLFunciones = new Funciones.Comunes();

            try
            {
                #region Variables y objetos                

                SAPbouiCOM.Folder oFolder1 = (SAPbouiCOM.Folder)oFormReception.Items.Item("Folder1").Specific;

                SAPbouiCOM.ButtonCombo cbEvMa = (SAPbouiCOM.ButtonCombo)oFormReception.Items.Item("cbEvMa").Specific;

                #endregion

                #region Centra en pantalla formulario

                oFormReception.Left = (sboapp.Desktop.Width - oFormReception.Width) / 2;
                oFormReception.Top = (sboapp.Desktop.Height - oFormReception.Height) / 4;

                #endregion

                #region Combo Masivamente

                cbEvMa.ValidValues.Add("10", "Acuse de recibo (DIAN)");
                cbEvMa.ValidValues.Add("12", "Recibo del bien y/o prestación del servicio");
                cbEvMa.ValidValues.Add("02", "Aceptación expresa (DIAN)");

                #endregion

                oFormReception.Visible = true;
                oFormReception.Refresh();
                oFolder1.Select();

            }
            catch (Exception e)
            {
                DLLFunciones.sendErrorMessage(sboapp, e);

            }
        }     

        public void DescargaXML_PDF_TFHKA(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Form _oFormVisorRepcecion)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {
                int iReprocesar = DllFunciones.sendMessageBoxY_N(_sboapp, "Se descargan los adjuntos de los documentos de compras sincronizados, ¿ Desea continuar ?");

                if (iReprocesar == 1)
                {
                    #region Consulta URL

                    string sGetModo = null;
                    string sURLRecepcion = null;
                    string sModo = null;
                    string sRutaXML = null;
                    string sRutaPDF = null;
                    string sProtocoloComunicacion = null;
                    string sTokenEmpresa = null;
                    string sTokenPassword = null;
                    string sGetDocumentsDownload = null;
                    string sSyncAttachmentOrigin = null;

                    SAPbobsCOM.Recordset oConsultarGetModo = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    sGetModo = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetModoandURL");

                    sGetModo = sGetModo.Replace("%Estado%", "\"U_BO_Status\" = 'Y'").Replace("%DocEntry%", " ");

                    oConsultarGetModo.DoQuery(sGetModo);

                    sURLRecepcion = Convert.ToString(oConsultarGetModo.Fields.Item("URLTFHKA").Value.ToString()) + "/ReceptorWS.svc?wsdl";
                    sModo = Convert.ToString(oConsultarGetModo.Fields.Item("Modo").Value.ToString());

                    sProtocoloComunicacion = Convert.ToString(oConsultarGetModo.Fields.Item("ProtocoloComunicacion").Value.ToString());
                    sTokenEmpresa = Convert.ToString(oConsultarGetModo.Fields.Item("TokenEmpresa").Value.ToString());
                    sTokenPassword = Convert.ToString(oConsultarGetModo.Fields.Item("TokenPassword").Value.ToString());
                    sRutaXML = Convert.ToString(oConsultarGetModo.Fields.Item("RutaXML").Value.ToString());
                    sRutaPDF = Convert.ToString(oConsultarGetModo.Fields.Item("RutaPDF").Value.ToString());

                    DllFunciones.liberarObjetos(oConsultarGetModo);

                    #endregion

                    #region Instanciacion parametros TFHKA

                    if (sProtocoloComunicacion == "HTTP")
                    {
                        BasicHttpBinding port = new BasicHttpBinding();
                    }
                    else if (sProtocoloComunicacion == "HTTPS")
                    {
                        BasicHttpsBinding port = new BasicHttpsBinding();
                    }

                    port.MaxBufferPoolSize = Int32.MaxValue;
                    port.MaxBufferSize = Int32.MaxValue;
                    port.MaxReceivedMessageSize = Int32.MaxValue;
                    port.ReaderQuotas.MaxStringContentLength = Int32.MaxValue;
                    port.SendTimeout = TimeSpan.FromMinutes(2);
                    port.ReceiveTimeout = TimeSpan.FromMinutes(2);

                    if (sProtocoloComunicacion == "HTTPS")
                    {
                        port.Security.Mode = BasicHttpSecurityMode.Transport;
                    }

                    EndpointAddress endPointEmision = new EndpointAddress(sURLRecepcion); //URL

                    Recepcion21WS.ReceptorWSClient serviceClienTFHKAReception;
                    serviceClienTFHKAReception = new Recepcion21WS.ReceptorWSClient(port, endPointEmision);

                    #endregion

                    #region Parametros generales Reporte

                    ReceptorRequestGeneral ParametrosDownload = new ReceptorRequestGeneral();

                    ParametrosDownload.tokenEmpresa = sTokenEmpresa;
                    ParametrosDownload.tokenPassword = sTokenPassword;

                    #endregion

                    #region Consulta documentos a descargar

                    SAPbobsCOM.Recordset oRsDocumentsDownload = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    SAPbobsCOM.Recordset oRsSyncDocument = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    sGetDocumentsDownload = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "sGetDocumentsDownload");
                    sSyncAttachmentOrigin = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "SyncAttachment");
                    
                    oRsDocumentsDownload.DoQuery(sGetDocumentsDownload);



                    #endregion

                    #region Descarga Documentos                     

                    if (oRsDocumentsDownload.RecordCount > 0)
                    {
                        Recepcion21WS.FileDownloadResponse DownloadXML;
                        Recepcion21WS.FileDownloadResponse DownloadPDF;

                        DllFunciones.ProgressBar(_oCompany, _sboapp, oRsDocumentsDownload.RecordCount + 1, 1, "Sincronizacion Adjuntos de los documentos, por favor espere...");

                        #region Descarga documentos

                        oRsDocumentsDownload.MoveFirst();

                        do
                        {
                            #region Parametros Complementarios Reporte

                            ParametrosDownload.identificadorEmisor = Convert.ToString(oRsDocumentsDownload.Fields.Item("U_BOTNI").Value.ToString());
                            ParametrosDownload.numeroDocumento = Convert.ToString(oRsDocumentsDownload.Fields.Item("U_BOTNF").Value.ToString());
                            ParametrosDownload.tipoIdentificacionemisor = Convert.ToString(oRsDocumentsDownload.Fields.Item("U_BOTTI").Value.ToString());

                            #endregion

                            #region Descarga XML

                            DownloadXML = serviceClienTFHKAReception.DescargarXML(ParametrosDownload);

                            string sRutaArchivoXML = sRutaXML + "\\" + Convert.ToString(oRsDocumentsDownload.Fields.Item("U_BOTNI").Value.ToString()) + "_" + Convert.ToString(oRsDocumentsDownload.Fields.Item("U_BOTNF").Value.ToString()) + ".xml";

                            if (DownloadXML.codigo == 200)
                            {
                                File.WriteAllBytes(sRutaArchivoXML, Convert.FromBase64String(DownloadXML.archivo));

                                string sSyncAttachmentCopy = sSyncAttachmentOrigin.Replace("%prmTipoAdjunto%", "XML").Replace("%pmrFile%", sRutaArchivoXML).Replace("%prmNIT%", Convert.ToString(oRsDocumentsDownload.Fields.Item("U_BOTNI").Value.ToString())).Replace("%prmNF%", Convert.ToString(oRsDocumentsDownload.Fields.Item("U_BOTNF").Value.ToString())).Replace("%CRWS%", DownloadXML.codigo.ToString());

                                oRsSyncDocument.DoQuery(sSyncAttachmentCopy);

                            }
                            else
                            {
                                string sSyncAttachmentCopy = sSyncAttachmentOrigin.Replace("%prmTipoAdjunto%", "XML").Replace("%pmrFile%", "").Replace("%prmNIT%", Convert.ToString(oRsDocumentsDownload.Fields.Item("U_BOTNI").Value.ToString())).Replace("%prmNF%", Convert.ToString(oRsDocumentsDownload.Fields.Item("U_BOTNF").Value.ToString())).Replace("%CRWS%", DownloadXML.codigo.ToString());

                                oRsSyncDocument.DoQuery(sSyncAttachmentCopy);

                            }

                            #endregion

                            #region Descarga PDF

                            DownloadPDF = serviceClienTFHKAReception.DescargarRepGrafica(ParametrosDownload);

                            string sRutaArchivoPDF = sRutaPDF + "\\" + Convert.ToString(oRsDocumentsDownload.Fields.Item("U_BOTNI").Value.ToString()) + "_" + Convert.ToString(oRsDocumentsDownload.Fields.Item("U_BOTNF").Value.ToString()) + ".pdf";

                            if (DownloadPDF.codigo == 200)
                            {
                                File.WriteAllBytes(sRutaArchivoPDF, Convert.FromBase64String(DownloadPDF.archivo));

                                string sSyncAttachmentCopy = sSyncAttachmentOrigin.Replace("%prmTipoAdjunto%", "PDF").Replace("%pmrFile%", sRutaArchivoPDF).Replace("%prmNIT%", Convert.ToString(oRsDocumentsDownload.Fields.Item("U_BOTNI").Value.ToString())).Replace("%prmNF%", Convert.ToString(oRsDocumentsDownload.Fields.Item("U_BOTNF").Value.ToString())).Replace("%CRWS%", DownloadPDF.codigo.ToString());

                                oRsSyncDocument.DoQuery(sSyncAttachmentCopy);
                            }
                            else
                            {
                                string sSyncAttachmentCopy = sSyncAttachmentOrigin.Replace("%prmTipoAdjunto%", "PDF").Replace("%pmrFile%", "").Replace("%prmNIT%", Convert.ToString(oRsDocumentsDownload.Fields.Item("U_BOTNI").Value.ToString())).Replace("%prmNF%", Convert.ToString(oRsDocumentsDownload.Fields.Item("U_BOTNF").Value.ToString())).Replace("%CRWS%", DownloadPDF.codigo.ToString());

                                oRsSyncDocument.DoQuery(sSyncAttachmentCopy);

                            }

                            #endregion

                            oRsDocumentsDownload.MoveNext();

                            DllFunciones.ProgressBar(_oCompany, _sboapp, oRsDocumentsDownload.RecordCount + 1, 1, "Sincronizando Adjuntos de los documentos, por favor espere...");

                        } while (oRsDocumentsDownload.EoF == false);

                        #endregion

                        DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Sincronizacion Finalizada.");
                    }

                    #endregion

                    #region Carga Infortmacion en la Matrix

                    #region Variabl1es y Objetos

                    string sPath;
                    string sInvoices = null;
                    string sCreditMemo = null;
                    string sDebitMemo = null;
                    int CantidadRegistos = 0;

                    string sPathImages = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\BOElectronicReception\\Images\\";

                    #endregion

                    sPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

                    SAPbouiCOM.EditText oFI = (SAPbouiCOM.EditText)_oFormVisorRepcecion.Items.Item("txtFI").Specific;
                    SAPbouiCOM.EditText oFF = (SAPbouiCOM.EditText)_oFormVisorRepcecion.Items.Item("txtFF").Specific;
                    SAPbouiCOM.ComboBox oEstado = (SAPbouiCOM.ComboBox)_oFormVisorRepcecion.Items.Item("cboStado").Specific;

                    #region Consulta de documentos facturas, notas debito y notas credito a mostrar en matrix

                    SAPbouiCOM.Matrix oMatrixInvoice = (Matrix)_oFormVisorRepcecion.Items.Item("MtxOPCH").Specific;
                    SAPbouiCOM.Matrix oMatrixCreditMemo = (Matrix)_oFormVisorRepcecion.Items.Item("MtxORPC").Specific;
                    SAPbouiCOM.Matrix oMatrixDebitMemo = (Matrix)_oFormVisorRepcecion.Items.Item("MtxOPCHD").Specific;

                    SAPbobsCOM.Recordset oRecorsetInvoices = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    SAPbobsCOM.Recordset oRecorsetCreditMemo = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    SAPbobsCOM.Recordset oRecorsetDebitMemo = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    SAPbouiCOM.DataTable oTableInvoices = _oFormVisorRepcecion.DataSources.DataTables.Item("DT_Invoices");
                    SAPbouiCOM.DataTable oTableCreditMemo = _oFormVisorRepcecion.DataSources.DataTables.Item("DT_CreditMemo");
                    SAPbouiCOM.DataTable oTableDebitMemo = _oFormVisorRepcecion.DataSources.DataTables.Item("DT_DebitMemo");

                    sInvoices = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetInvoices");
                    sInvoices = sInvoices.Replace("%PathImages%", sPathImages);

                    sCreditMemo = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetCreditMemo");
                    sInvoices = sInvoices.Replace("%PathImages%", sPathImages);

                    sDebitMemo = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetDebitMemo");
                    sInvoices = sInvoices.Replace("%PathImages%", sPathImages);


                    if (string.IsNullOrEmpty(oFI.Value))
                    {
                        sInvoices = sInvoices.Replace("%FI%", "20190101");
                    }
                    else
                    {
                        sInvoices = sInvoices.Replace("%FI%", oFI.Value);
                    }


                    if (string.IsNullOrEmpty(oFF.Value))
                    {
                        sInvoices = sInvoices.Replace("%FF%", "20301231");
                    }
                    else
                    {
                        sInvoices = sInvoices.Replace("%FF%", oFF.Value);
                    }

                    if (oEstado.Value == "-")
                    {
                        sInvoices = sInvoices.Replace("%Estado%", "");
                    }
                    else
                    {
                        sInvoices = sInvoices.Replace("%Estado%", "AND \"U_BOTSDC\" = '"+oEstado.Value+"' " );
                    }


                    oRecorsetInvoices.DoQuery(sInvoices);
                    oRecorsetCreditMemo.DoQuery(sCreditMemo);
                    oRecorsetDebitMemo.DoQuery(sDebitMemo);
                    
                    oTableInvoices.ExecuteQuery(sInvoices);
                    oTableCreditMemo.ExecuteQuery(sCreditMemo);
                    oTableDebitMemo.ExecuteQuery(sDebitMemo);

                    #endregion

                    CantidadRegistos = oRecorsetInvoices.RecordCount + oRecorsetCreditMemo.RecordCount + oRecorsetDebitMemo.RecordCount;

                    if (CantidadRegistos != 0)
                    {
                        #region Carga datos Matrix Facturas

                        if (oRecorsetInvoices.RecordCount > 0)
                        {
                            oMatrixInvoice.Clear();

                            oMatrixInvoice.Columns.Item("#").DataBind.Bind("DT_Invoices", "#");
                            oMatrixInvoice.Columns.Item("Col_0").DataBind.Bind("DT_Invoices", "Estado");
                            oMatrixInvoice.Columns.Item("Col_11").DataBind.Bind("DT_Invoices", "Estado_SAP");
                            oMatrixInvoice.Columns.Item("Col_9").DataBind.Bind("DT_Invoices", "DocEntry");
                            oMatrixInvoice.Columns.Item("Col_1").DataBind.Bind("DT_Invoices", "Num_Fac_Pro");
                            oMatrixInvoice.Columns.Item("Col_16").DataBind.Bind("DT_Invoices", "Serie_Numeracion");
                            oMatrixInvoice.Columns.Item("Col_2").DataBind.Bind("DT_Invoices", "Codigo_cliente");
                            oMatrixInvoice.Columns.Item("Col_3").DataBind.Bind("DT_Invoices", "Nombre_del_Cliente");
                            oMatrixInvoice.Columns.Item("Col_4").DataBind.Bind("DT_Invoices", "Fecha_Documento");
                            oMatrixInvoice.Columns.Item("Col_5").DataBind.Bind("DT_Invoices", "Fecha_vencimiento");
                            oMatrixInvoice.Columns.Item("Col_24").DataBind.Bind("DT_Invoices", "Condicion_de_Pago");
                            oMatrixInvoice.Columns.Item("Col_6").DataBind.Bind("DT_Invoices", "Total_documento");
                            oMatrixInvoice.Columns.Item("Col_17").DataBind.Bind("DT_Invoices", "CUFE");
                            oMatrixInvoice.Columns.Item("Col_18").DataBind.Bind("DT_Invoices", "Fecha_emision");
                            oMatrixInvoice.Columns.Item("Col_10").DataBind.Bind("DT_Invoices", "Hora_emision");
                            oMatrixInvoice.Columns.Item("Col_8").DataBind.Bind("DT_Invoices", "Fecha_recepcion");
                            oMatrixInvoice.Columns.Item("Col_19").Visible = false;
                            oMatrixInvoice.Columns.Item("Col_23").Visible = false;
                            oMatrixInvoice.Columns.Item("Col_20").DataBind.Bind("DT_Invoices", "DescargaXML");
                            oMatrixInvoice.Columns.Item("Col_12").DataBind.Bind("DT_Invoices", "RutaXML");
                            oMatrixInvoice.Columns.Item("Col_12").Visible = false;
                            oMatrixInvoice.Columns.Item("Col_21").DataBind.Bind("DT_Invoices", "DescargaPDF");
                            oMatrixInvoice.Columns.Item("Col_13").DataBind.Bind("DT_Invoices", "RutaPDF");
                            oMatrixInvoice.Columns.Item("Col_13").Visible = false;
                            oMatrixInvoice.Columns.Item("Col_22").DataBind.Bind("DT_Invoices", "ImageAceptar");
                            oMatrixInvoice.Columns.Item("Col_7").DataBind.Bind("DT_Invoices", "ImageCancelar");
                            oMatrixInvoice.Columns.Item("Col_14").DataBind.Bind("DT_Invoices", "TipoIdentificacionEmisor");
                            oMatrixInvoice.Columns.Item("Col_14").Visible = false;


                            oMatrixInvoice.LoadFromDataSource();

                            oMatrixInvoice.AutoResizeColumns();

                        }

                        #endregion

                        #region Carga datos Matrix Notas credito

                        if (oRecorsetCreditMemo.RecordCount > 0)
                        {
                            oMatrixCreditMemo.Clear();

                            oMatrixCreditMemo.Columns.Item("#").DataBind.Bind("DT_Invoices", "#");
                            oMatrixCreditMemo.Columns.Item("Col_0").DataBind.Bind("DT_Invoices", "Estado");
                            oMatrixCreditMemo.Columns.Item("Col_9").DataBind.Bind("DT_Invoices", "DocEntry");
                            oMatrixCreditMemo.Columns.Item("Col_1").DataBind.Bind("DT_Invoices", "Num_Fac_Pro");
                            //oMatrixCreditMemo.Columns.Item("Col_25").DataBind.Bind("DT_Invoices", "Num_Fac_Preeli");
                            //oMatrixCreditMemo.Columns.Item("Col_26").DataBind.Bind("DT_Invoices", "Num_Fac_SAP");
                            oMatrixCreditMemo.Columns.Item("Col_16").DataBind.Bind("DT_Invoices", "Serie_Numeracion");
                            oMatrixCreditMemo.Columns.Item("Col_2").DataBind.Bind("DT_Invoices", "Codigo_cliente");
                            oMatrixCreditMemo.Columns.Item("Col_3").DataBind.Bind("DT_Invoices", "Nombre_del_Cliente");
                            oMatrixCreditMemo.Columns.Item("Col_4").DataBind.Bind("DT_Invoices", "Fecha_Documento");
                            oMatrixCreditMemo.Columns.Item("Col_5").DataBind.Bind("DT_Invoices", "Fecha_vencimiento");
                            oMatrixCreditMemo.Columns.Item("Col_24").DataBind.Bind("DT_Invoices", "Condicion_Pago");
                            oMatrixCreditMemo.Columns.Item("Col_6").DataBind.Bind("DT_Invoices", "Total_documento");
                            oMatrixCreditMemo.Columns.Item("Col_17").DataBind.Bind("DT_Invoices", "CUFE");
                            oMatrixCreditMemo.Columns.Item("Col_18").DataBind.Bind("DT_Invoices", "Fecha_emision");
                            oMatrixCreditMemo.Columns.Item("Col_10").DataBind.Bind("DT_Invoices", "Hora_emision");
                            oMatrixCreditMemo.Columns.Item("Col_8").DataBind.Bind("DT_Invoices", "Fecha_recepcion");

                            oMatrixCreditMemo.LoadFromDataSource();

                            oMatrixCreditMemo.AutoResizeColumns();

                        }
                        #endregion

                        #region Carga datos Matrix Notas Debito

                        if (oRecorsetDebitMemo.RecordCount > 0)
                        {
                            oMatrixDebitMemo.Clear();

                            oMatrixDebitMemo.Columns.Item("#").DataBind.Bind("DT_Invoices", "#");
                            oMatrixDebitMemo.Columns.Item("Col_0").DataBind.Bind("DT_Invoices", "Estado");
                            oMatrixDebitMemo.Columns.Item("Col_9").DataBind.Bind("DT_Invoices", "DocEntry");
                            oMatrixDebitMemo.Columns.Item("Col_1").DataBind.Bind("DT_Invoices", "Num_Fac_Pro");
                            //oMatrixDebitMemo.Columns.Item("Col_25").DataBind.Bind("DT_Invoices", "Num_Fac_Preeli");
                            //oMatrixDebitMemo.Columns.Item("Col_26").DataBind.Bind("DT_Invoices", "Num_Fac_SAP");
                            oMatrixDebitMemo.Columns.Item("Col_16").DataBind.Bind("DT_Invoices", "Serie_Numeracion");
                            oMatrixDebitMemo.Columns.Item("Col_2").DataBind.Bind("DT_Invoices", "Codigo_cliente");
                            oMatrixDebitMemo.Columns.Item("Col_3").DataBind.Bind("DT_Invoices", "Nombre_del_Cliente");
                            oMatrixDebitMemo.Columns.Item("Col_4").DataBind.Bind("DT_Invoices", "Fecha_Documento");
                            oMatrixDebitMemo.Columns.Item("Col_5").DataBind.Bind("DT_Invoices", "Fecha_vencimiento");
                            oMatrixDebitMemo.Columns.Item("Col_24").DataBind.Bind("DT_Invoices", "Condicion_Pago");
                            oMatrixDebitMemo.Columns.Item("Col_6").DataBind.Bind("DT_Invoices", "Total_documento");
                            oMatrixDebitMemo.Columns.Item("Col_17").DataBind.Bind("DT_Invoices", "CUFE");
                            oMatrixDebitMemo.Columns.Item("Col_18").DataBind.Bind("DT_Invoices", "Fecha_emision");
                            oMatrixDebitMemo.Columns.Item("Col_10").DataBind.Bind("DT_Invoices", "Hora_emision");
                            oMatrixDebitMemo.Columns.Item("Col_8").DataBind.Bind("DT_Invoices", "Fecha_recepcion");

                            oMatrixDebitMemo.LoadFromDataSource();

                            oMatrixDebitMemo.AutoResizeColumns();

                        }
                        #endregion

                    }
                    else
                    {
                        DllFunciones.sendMessageBox(_sboapp, "No se encontraron documentos");
                    }

                    DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Documentos sicronizados correctamente.");

                    #endregion

                    #region Liberacion de Objetos

                    DllFunciones.liberarObjetos(oRsDocumentsDownload);
                    DllFunciones.liberarObjetos(oRsSyncDocument);

                    #endregion
                }
            }
            catch (Exception)
            {

                throw;
            }

        }

        public string GetElementXML_DIAN(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _oCompany, string Elemento_XML_DIAN, string sRutaXML, string U_BOTNI, string U_BOTNF )
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {
                string sRutaArchivoXML = sRutaXML + "\\" + U_BOTNI + "_" + U_BOTNF + ".xml";

                if (File.Exists(sRutaArchivoXML))
                {
                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.Load(sRutaArchivoXML);

                    XmlNamespaceManager nsMain = new XmlNamespaceManager(xmlDoc.NameTable);
                    nsMain.AddNamespace("cbc", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2");
                    nsMain.AddNamespace("cac", "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2");
                    
                    if (Elemento_XML_DIAN == "OrderReference")
                    {
                        #region Retorna el Elemento OrderReference

                        // Paso 1: obtener el contenido del CDATA
                        XmlNode cdataNode = xmlDoc.SelectSingleNode("//cac:Attachment/cac:ExternalReference/cbc:Description", nsMain);
                        if (cdataNode != null)
                        {
                            string embeddedXml = cdataNode.InnerText;

                            // Paso 2: cargar el XML embebido
                            XmlDocument embeddedDoc = new XmlDocument();
                            embeddedDoc.LoadXml(embeddedXml);

                            XmlNamespaceManager ns = new XmlNamespaceManager(embeddedDoc.NameTable);
                            ns.AddNamespace("cbc", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2");
                            ns.AddNamespace("cac", "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2");

                            // Paso 3: extraer el ID dentro de PaymentMeans
                            XmlNode idNode = embeddedDoc.SelectSingleNode("//cac:OrderReference/cbc:ID", ns);
                            if (idNode != null)
                            {
                                return idNode.InnerText;
                            }
                            else
                            {
                                Console.WriteLine("Nodo <cbc:ID> dentro de <cac:PaymentMeans> no encontrado.");
                                return "";
                            }
                        }
                        else
                        {
                            Console.WriteLine("Nodo <cbc:Description> con XML embebido no encontrado.");
                            return "";
                        }

                        #endregion
                    }

                    if (Elemento_XML_DIAN == "PaymentMeans_ID")
                    {
                        #region Retorna el Elemento OrderReference

                        // Paso 1: obtener el contenido del CDATA
                        XmlNode cdataNode = xmlDoc.SelectSingleNode("//cac:Attachment/cac:ExternalReference/cbc:Description", nsMain);
                        if (cdataNode != null)
                        {
                            string embeddedXml = cdataNode.InnerText;

                            // Paso 2: cargar el XML embebido
                            XmlDocument embeddedDoc = new XmlDocument();
                            embeddedDoc.LoadXml(embeddedXml);

                            XmlNamespaceManager ns = new XmlNamespaceManager(embeddedDoc.NameTable);
                            ns.AddNamespace("cbc", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2");
                            ns.AddNamespace("cac", "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2");

                            // Paso 3: extraer el ID dentro de PaymentMeans
                            XmlNode idNode = embeddedDoc.SelectSingleNode("//cac:PaymentMeans/cbc:ID", ns);
                            if (idNode != null)
                            {                                
                                return idNode.InnerText;
                            }
                            else
                            {
                                Console.WriteLine("Nodo <cbc:ID> dentro de <cac:PaymentMeans> no encontrado.");
                                return "";
                            }
                        }
                        else
                        {
                            Console.WriteLine("Nodo <cbc:Description> con XML embebido no encontrado.");
                            return "";
                        }

                        #endregion
                    }
                    else
                    {
                        return "";
                    }
                }
                else
                {
                    return "";
                }


            }
            catch (Exception)
            {
                throw;
            }

        }
        
        public void MatrixOpenFile(SAPbouiCOM.Application sboappElectronicReception, SAPbobsCOM.Company oCompanyElectronicReception, ItemEvent pVal,string ColID)
        {
            Funciones.Comunes DLLFunciones = new Funciones.Comunes();
            try
            {
                SAPbouiCOM.Form oForm = null;
                bool ExistForm = false;

                for (int i = 0; i < sboappElectronicReception.Forms.Count; i++)
                {
                    if (sboappElectronicReception.Forms.Item(i).UniqueID == "BOTVDR")
                    {
                        oForm = sboappElectronicReception.Forms.Item("BOTVDR");
                        ExistForm = true;
                    }
                }


                if (ExistForm)
                {
                    SAPbobsCOM.Recordset oRsProveedorTecnologico = (SAPbobsCOM.Recordset)oCompanyElectronicReception.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    string sGetProveedorTecnologico = null;

                    sGetProveedorTecnologico = DLLFunciones.GetStringXMLDocument(oCompanyElectronicReception, "eBilling", "eBilling", "GetProveedorTecnologico");
                    oRsProveedorTecnologico.DoQuery(sGetProveedorTecnologico);

                    if (oRsProveedorTecnologico.RecordCount > 0)
                    {
                        if (oRsProveedorTecnologico.Fields.Item("ProveedorTecnologico").Value.ToString() == "TFHKA")
                        {
                            MatrixOpenFile_TFHKA(sboappElectronicReception, oCompanyElectronicReception, oForm, pVal, ColID);

                        }
                        else if (oRsProveedorTecnologico.Fields.Item("ProveedorTecnologico").Value.ToString() == "FBE")
                        {
                            MatrixOpenFile_TFHKA(sboappElectronicReception, oCompanyElectronicReception, oForm, pVal, ColID);
                        }
                    }
                }
                else
                {

                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        public void ChagueStatusDocument(SAPbouiCOM.Application sboappElectronicReception, SAPbobsCOM.Company oCompanyElectronicReception, ItemEvent pVal, string ColID)
        {
            Funciones.Comunes DLLFunciones = new Funciones.Comunes();
            try
            {
                SAPbouiCOM.Form oForm = null;
                bool ExistForm = false;

                for (int i = 0; i < sboappElectronicReception.Forms.Count; i++)
                {
                    if (sboappElectronicReception.Forms.Item(i).UniqueID == "BOTVDR")
                    {
                        oForm = sboappElectronicReception.Forms.Item("BOTVDR");
                        ExistForm = true;
                    }
                }


                if (ExistForm)
                {
                    SAPbobsCOM.Recordset oRsProveedorTecnologico = (SAPbobsCOM.Recordset)oCompanyElectronicReception.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    string sGetProveedorTecnologico = null;

                    sGetProveedorTecnologico = DLLFunciones.GetStringXMLDocument(oCompanyElectronicReception, "eBilling", "eBilling", "GetProveedorTecnologico");
                    oRsProveedorTecnologico.DoQuery(sGetProveedorTecnologico);

                    if (oRsProveedorTecnologico.RecordCount > 0)
                    {
                        if (oRsProveedorTecnologico.Fields.Item("ProveedorTecnologico").Value.ToString() == "TFHKA")
                        {
                            ChagueStatusDocument_TFHKA(sboappElectronicReception, oCompanyElectronicReception, oForm, pVal, ColID);

                        }
                        else if (oRsProveedorTecnologico.Fields.Item("ProveedorTecnologico").Value.ToString() == "FBE")
                        {
                            ChagueStatusDocument_TFHKA(sboappElectronicReception, oCompanyElectronicReception, oForm, pVal, ColID);
                        }
                    }
                }
                else
                {

                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        public void ChagueDocumentStatusBulk(SAPbouiCOM.Application sboappElectronicReception, SAPbobsCOM.Company oCompanyElectronicReception, ItemEvent pVal, string ColID)
        {
            Funciones.Comunes DLLFunciones = new Funciones.Comunes();
            try
            {
                SAPbouiCOM.Form oForm = null;
                bool ExistForm = false;

                for (int i = 0; i < sboappElectronicReception.Forms.Count; i++)
                {
                    if (sboappElectronicReception.Forms.Item(i).UniqueID == "BOTVDR")
                    {
                        oForm = sboappElectronicReception.Forms.Item("BOTVDR");
                        ExistForm = true;
                    }
                }


                if (ExistForm)
                {
                    SAPbobsCOM.Recordset oRsProveedorTecnologico = (SAPbobsCOM.Recordset)oCompanyElectronicReception.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    string sGetProveedorTecnologico = null;

                    sGetProveedorTecnologico = DLLFunciones.GetStringXMLDocument(oCompanyElectronicReception, "eBilling", "eBilling", "GetProveedorTecnologico");
                    oRsProveedorTecnologico.DoQuery(sGetProveedorTecnologico);

                    if (oRsProveedorTecnologico.RecordCount > 0)
                    {
                        if (oRsProveedorTecnologico.Fields.Item("ProveedorTecnologico").Value.ToString() == "TFHKA")
                        {
                            ChagueDocumentStatusBulk_TFHKA(sboappElectronicReception, oCompanyElectronicReception, oForm, pVal, ColID);
                        }
                        else if (oRsProveedorTecnologico.Fields.Item("ProveedorTecnologico").Value.ToString() == "FBE")
                        {
                            ChagueDocumentStatusBulk_TFHKA(sboappElectronicReception, oCompanyElectronicReception, oForm, pVal, ColID);
                        }
                    }
                }
                else
                {

                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        public void LoadMatrixReception(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Form _oFormVisorRepcecion)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            SAPbouiCOM.EditText oFI = (SAPbouiCOM.EditText)_oFormVisorRepcecion.Items.Item("txtFI").Specific;
            SAPbouiCOM.EditText oFF = (SAPbouiCOM.EditText)_oFormVisorRepcecion.Items.Item("txtFF").Specific;
            SAPbouiCOM.EditText otxtND = (SAPbouiCOM.EditText)_oFormVisorRepcecion.Items.Item("txtND").Specific;
            SAPbouiCOM.ComboBox oEstado = (SAPbouiCOM.ComboBox)_oFormVisorRepcecion.Items.Item("cboStado").Specific;

            if (string.IsNullOrEmpty(oFI.Value))
            {
                DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10026", _sboapp, null) );
            }
            else if (string.IsNullOrEmpty(oFF.Value))
            {
                DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10027", _sboapp, null));
            }
            else
            {
                #region Carga informacion en la Matrix

                #region Variables y Objetos

                string sPath;
                string sInvoices = null;
                string sCreditMemo = null;
                string sDebitMemo = null;
                int CantidadRegistos = 0;

                string sPathImages = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\BOElectronicReception\\Images\\";

                #endregion

                sPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

                #region Consulta de documentos facturas, notas debito y notas credito a mostrar en matrix

                SAPbouiCOM.Matrix oMatrixInvoice = (Matrix)_oFormVisorRepcecion.Items.Item("MtxOPCH").Specific;
                SAPbouiCOM.Matrix oMatrixCreditMemo = (Matrix)_oFormVisorRepcecion.Items.Item("MtxORPC").Specific;
                SAPbouiCOM.Matrix oMatrixDebitMemo = (Matrix)_oFormVisorRepcecion.Items.Item("MtxOPCHD").Specific;

                SAPbobsCOM.Recordset oRecorsetInvoices = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oRecorsetCreditMemo = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oRecorsetDebitMemo = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                SAPbouiCOM.DataTable oTableInvoices = _oFormVisorRepcecion.DataSources.DataTables.Item("DT_Invoices");
                SAPbouiCOM.DataTable oTableCreditMemo = _oFormVisorRepcecion.DataSources.DataTables.Item("DT_CreditMemo");
                SAPbouiCOM.DataTable oTableDebitMemo = _oFormVisorRepcecion.DataSources.DataTables.Item("DT_DebitMemo");

                sInvoices = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetInvoices");
                sInvoices = sInvoices.Replace("%PathImages%", sPathImages).Replace("%FI%", oFI.Value.ToString()).Replace("%FF%",oFF.Value.ToString());

                sCreditMemo = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetCreditMemo");
                sInvoices = sInvoices.Replace("%PathImages%", sPathImages).Replace("%FI%", oFI.Value.ToString()).Replace("%FF%", oFF.Value.ToString());

                sDebitMemo = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetDebitMemo");
                sInvoices = sInvoices.Replace("%PathImages%", sPathImages).Replace("%FI%", oFI.Value.ToString()).Replace("%FF%", oFF.Value.ToString());

                if (oEstado.Value == "-" || oEstado.Value == "")
                {
                    sInvoices = sInvoices.Replace("%Estado%", "");
                }
                else
                {
                    sInvoices = sInvoices.Replace("%Estado%", "AND \"U_BOTSDC\" = '" + oEstado.Value + "' ");
                }

                if (otxtND.Value == "")
                {
                    sInvoices = sInvoices.Replace("%NumeroDocumento%", "");
                }
                else
                {
                    sInvoices = sInvoices.Replace("%NumeroDocumento%", "AND \"U_BOTNF\" LIKE '%" + otxtND.Value + "%' ");
                }

                oRecorsetInvoices.DoQuery(sInvoices);

                oRecorsetCreditMemo.DoQuery(sCreditMemo);

                oRecorsetDebitMemo.DoQuery(sDebitMemo);

                oTableInvoices.ExecuteQuery(sInvoices);
                oTableCreditMemo.ExecuteQuery(sCreditMemo);
                oTableDebitMemo.ExecuteQuery(sDebitMemo);

                #endregion

                CantidadRegistos = oRecorsetInvoices.RecordCount + oRecorsetCreditMemo.RecordCount + oRecorsetDebitMemo.RecordCount;

                if (CantidadRegistos != 0)
                {
                    DllFunciones.StatusBar(_sboapp,BoStatusBarMessageType.smt_Warning, MessageSystemAddOn("FER10036", _sboapp, null ));

                    _oFormVisorRepcecion.Freeze(true);

                    #region Carga datos Matrix Facturas

                    if (oRecorsetInvoices.RecordCount > 0)
                    {
                        oMatrixInvoice.Clear();

                        oMatrixInvoice.Columns.Item("#").DataBind.Bind("DT_Invoices", "#");
                        oMatrixInvoice.Columns.Item("Col_0").DataBind.Bind("DT_Invoices", "Estado");
                        oMatrixInvoice.Columns.Item("Col_11").DataBind.Bind("DT_Invoices", "Estado_SAP");
                        oMatrixInvoice.Columns.Item("Col_9").DataBind.Bind("DT_Invoices", "DocEntry");
                        oMatrixInvoice.Columns.Item("Col_1").DataBind.Bind("DT_Invoices", "Num_Fac_Pro");
                        oMatrixInvoice.Columns.Item("Col_16").DataBind.Bind("DT_Invoices", "Serie_Numeracion");
                        oMatrixInvoice.Columns.Item("Col_2").DataBind.Bind("DT_Invoices", "Codigo_cliente");
                        oMatrixInvoice.Columns.Item("Col_3").DataBind.Bind("DT_Invoices", "Nombre_del_Cliente");
                        oMatrixInvoice.Columns.Item("Col_4").DataBind.Bind("DT_Invoices", "Fecha_Documento");
                        oMatrixInvoice.Columns.Item("Col_5").DataBind.Bind("DT_Invoices", "Fecha_vencimiento");
                        oMatrixInvoice.Columns.Item("Col_24").DataBind.Bind("DT_Invoices", "Condicion_de_Pago");
                        oMatrixInvoice.Columns.Item("Col_6").DataBind.Bind("DT_Invoices", "Total_documento");
                        oMatrixInvoice.Columns.Item("Col_17").DataBind.Bind("DT_Invoices", "CUFE");
                        oMatrixInvoice.Columns.Item("Col_18").DataBind.Bind("DT_Invoices", "Fecha_emision");
                        oMatrixInvoice.Columns.Item("Col_10").DataBind.Bind("DT_Invoices", "Hora_emision");
                        oMatrixInvoice.Columns.Item("Col_8").DataBind.Bind("DT_Invoices", "Fecha_recepcion");
                        oMatrixInvoice.Columns.Item("Col_19").Visible = false;
                        oMatrixInvoice.Columns.Item("Col_23").Visible = false;
                        oMatrixInvoice.Columns.Item("Col_20").DataBind.Bind("DT_Invoices", "DescargaXML");
                        oMatrixInvoice.Columns.Item("Col_12").DataBind.Bind("DT_Invoices", "RutaXML");
                        oMatrixInvoice.Columns.Item("Col_12").Visible = false;
                        oMatrixInvoice.Columns.Item("Col_21").DataBind.Bind("DT_Invoices", "DescargaPDF");
                        oMatrixInvoice.Columns.Item("Col_13").DataBind.Bind("DT_Invoices", "RutaPDF");
                        oMatrixInvoice.Columns.Item("Col_13").Visible = false;
                        oMatrixInvoice.Columns.Item("Col_22").DataBind.Bind("DT_Invoices", "ImageAceptar");
                        oMatrixInvoice.Columns.Item("Col_7").DataBind.Bind("DT_Invoices", "ImageCancelar");
                        oMatrixInvoice.Columns.Item("Col_14").DataBind.Bind("DT_Invoices", "TipoIdentificacionEmisor");
                        oMatrixInvoice.Columns.Item("Col_14").Visible = false;
                        oMatrixInvoice.Columns.Item("Col_15").DataBind.Bind("DT_Invoices", "ImageAcusar");
                        oMatrixInvoice.Columns.Item("Col_27").DataBind.Bind("DT_Invoices", "ImageRecibir");
                        oMatrixInvoice.Columns.Item("Col_28").DataBind.Bind("DT_Invoices", "Chk");
                        oMatrixInvoice.Columns.Item("Col_28").Editable = true;
                        oMatrixInvoice.Columns.Item("Col_29").DataBind.Bind("DT_Invoices", "PO");
                        oMatrixInvoice.Columns.Item("Col_29").Editable = false;


                        oMatrixInvoice.LoadFromDataSource();

                        oMatrixInvoice.AutoResizeColumns();


                        #region Valida si se activa el campo "Seleccionar"

                        int rowCount = oMatrixInvoice.RowCount;

                        for (int i = 1; i <= rowCount; i++)
                        {
                            // Obtener el valor de la columna de control (por ejemplo, "ColYN")
                            //string habilitado = oMatrixInvoice.Columns.Item("Col_24").Cells.Item(i).Specific.Value;
                            string sMedioPagoDIAN = ((SAPbouiCOM.EditText)(oMatrixInvoice.Columns.Item("Col_24").Cells.Item(i).Specific)).Value;

                            // Obtener la celda a habilitar/deshabilitar (por ejemplo, columna "ColCheck")
                            SAPbouiCOM.CheckBox txtSeleccionar = (SAPbouiCOM.CheckBox)oMatrixInvoice.Columns.Item("Col_28").Cells.Item(i).Specific;

                            if ( sMedioPagoDIAN == "Contado" )
                            {
                                txtSeleccionar.Item.Enabled = false;  // habilita la celda
                                oMatrixInvoice.CommonSetting.SetCellEditable(i, 1, false);
                            }
                            else
                            {
                                txtSeleccionar.Item.Enabled = true; // deshabilita la celda
                            }
                        }

                        

                        #endregion


                    }

                    #endregion

                    #region Carga datos Matrix Notas credito

                    if (oRecorsetCreditMemo.RecordCount > 0)
                    {
                        oMatrixCreditMemo.Clear();

                        oMatrixCreditMemo.Columns.Item("#").DataBind.Bind("DT_Invoices", "#");
                        oMatrixCreditMemo.Columns.Item("Col_0").DataBind.Bind("DT_Invoices", "Estado");
                        oMatrixCreditMemo.Columns.Item("Col_9").DataBind.Bind("DT_Invoices", "DocEntry");
                        oMatrixCreditMemo.Columns.Item("Col_1").DataBind.Bind("DT_Invoices", "Num_Fac_Pro");
                        //oMatrixCreditMemo.Columns.Item("Col_25").DataBind.Bind("DT_Invoices", "Num_Fac_Preeli");
                        //oMatrixCreditMemo.Columns.Item("Col_26").DataBind.Bind("DT_Invoices", "Num_Fac_SAP");
                        oMatrixCreditMemo.Columns.Item("Col_16").DataBind.Bind("DT_Invoices", "Serie_Numeracion");
                        oMatrixCreditMemo.Columns.Item("Col_2").DataBind.Bind("DT_Invoices", "Codigo_cliente");
                        oMatrixCreditMemo.Columns.Item("Col_3").DataBind.Bind("DT_Invoices", "Nombre_del_Cliente");
                        oMatrixCreditMemo.Columns.Item("Col_4").DataBind.Bind("DT_Invoices", "Fecha_Documento");
                        oMatrixCreditMemo.Columns.Item("Col_5").DataBind.Bind("DT_Invoices", "Fecha_vencimiento");
                        oMatrixCreditMemo.Columns.Item("Col_24").DataBind.Bind("DT_Invoices", "Condicion_Pago");
                        oMatrixCreditMemo.Columns.Item("Col_6").DataBind.Bind("DT_Invoices", "Total_documento");
                        oMatrixCreditMemo.Columns.Item("Col_17").DataBind.Bind("DT_Invoices", "CUFE");
                        oMatrixCreditMemo.Columns.Item("Col_18").DataBind.Bind("DT_Invoices", "Fecha_emision");
                        oMatrixCreditMemo.Columns.Item("Col_10").DataBind.Bind("DT_Invoices", "Hora_emision");
                        oMatrixCreditMemo.Columns.Item("Col_8").DataBind.Bind("DT_Invoices", "Fecha_recepcion");

                        oMatrixCreditMemo.LoadFromDataSource();

                        oMatrixCreditMemo.AutoResizeColumns();

                    }
                    #endregion

                    #region Carga datos Matrix Notas Debito

                    if (oRecorsetDebitMemo.RecordCount > 0)
                    {
                        oMatrixDebitMemo.Clear();

                        oMatrixDebitMemo.Columns.Item("#").DataBind.Bind("DT_Invoices", "#");
                        oMatrixDebitMemo.Columns.Item("Col_0").DataBind.Bind("DT_Invoices", "Estado");
                        oMatrixDebitMemo.Columns.Item("Col_9").DataBind.Bind("DT_Invoices", "DocEntry");
                        oMatrixDebitMemo.Columns.Item("Col_1").DataBind.Bind("DT_Invoices", "Num_Fac_Pro");
                        //oMatrixDebitMemo.Columns.Item("Col_25").DataBind.Bind("DT_Invoices", "Num_Fac_Preeli");
                        //oMatrixDebitMemo.Columns.Item("Col_26").DataBind.Bind("DT_Invoices", "Num_Fac_SAP");
                        oMatrixDebitMemo.Columns.Item("Col_16").DataBind.Bind("DT_Invoices", "Serie_Numeracion");
                        oMatrixDebitMemo.Columns.Item("Col_2").DataBind.Bind("DT_Invoices", "Codigo_cliente");
                        oMatrixDebitMemo.Columns.Item("Col_3").DataBind.Bind("DT_Invoices", "Nombre_del_Cliente");
                        oMatrixDebitMemo.Columns.Item("Col_4").DataBind.Bind("DT_Invoices", "Fecha_Documento");
                        oMatrixDebitMemo.Columns.Item("Col_5").DataBind.Bind("DT_Invoices", "Fecha_vencimiento");
                        oMatrixDebitMemo.Columns.Item("Col_24").DataBind.Bind("DT_Invoices", "Condicion_Pago");
                        oMatrixDebitMemo.Columns.Item("Col_6").DataBind.Bind("DT_Invoices", "Total_documento");
                        oMatrixDebitMemo.Columns.Item("Col_17").DataBind.Bind("DT_Invoices", "CUFE");
                        oMatrixDebitMemo.Columns.Item("Col_18").DataBind.Bind("DT_Invoices", "Fecha_emision");
                        oMatrixDebitMemo.Columns.Item("Col_10").DataBind.Bind("DT_Invoices", "Hora_emision");
                        oMatrixDebitMemo.Columns.Item("Col_8").DataBind.Bind("DT_Invoices", "Fecha_recepcion");

                        oMatrixDebitMemo.LoadFromDataSource();

                        oMatrixDebitMemo.AutoResizeColumns();

                    }
                    #endregion

                    _oFormVisorRepcecion.Freeze(false);

                }
                else
                {
                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10023", _sboapp, null));
                }

                DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, MessageSystemAddOn("FER10024", _sboapp, null));

                #endregion

                #region Liberar Objetos

                DllFunciones.liberarObjetos(oRecorsetCreditMemo);
                DllFunciones.liberarObjetos(oRecorsetDebitMemo);

                #endregion
            }

        }

        public void DownloadDocumentElectronicReceipt(SAPbouiCOM.Application sboappElectronicReception, SAPbobsCOM.Company oCompanyElectronicReception)
        {
            Funciones.Comunes DLLFunciones = new Funciones.Comunes();
            try
            {
                SAPbouiCOM.Form oForm = null;
                bool ExistForm = false;

                for (int i = 0; i < sboappElectronicReception.Forms.Count; i++)
                {
                    if (sboappElectronicReception.Forms.Item(i).UniqueID == "BOTVDR")
                    {
                        oForm = sboappElectronicReception.Forms.Item("BOTVDR");
                        ExistForm = true;
                    }
                }


                if (ExistForm)
                {
                    SAPbobsCOM.Recordset oRsProveedorTecnologico = (SAPbobsCOM.Recordset)oCompanyElectronicReception.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    string sGetProveedorTecnologico = null;

                    sGetProveedorTecnologico = DLLFunciones.GetStringXMLDocument(oCompanyElectronicReception, "eBilling", "eBilling", "GetProveedorTecnologico");
                    oRsProveedorTecnologico.DoQuery(sGetProveedorTecnologico);

                    if (oRsProveedorTecnologico.RecordCount > 0)
                    {
                        if (oRsProveedorTecnologico.Fields.Item("ProveedorTecnologico").Value.ToString() == "TFHKA")
                        {
                            DescargaDocumentosTFHKA(sboappElectronicReception, oCompanyElectronicReception, oForm);
                        }
                        else if (oRsProveedorTecnologico.Fields.Item("ProveedorTecnologico").Value.ToString() == "FBE")
                        {
                            DescargaDocumentosFBE(sboappElectronicReception, oCompanyElectronicReception, oForm);
                        }                        
                    }
                }
                else
                {

                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        public void DownloadDocumentXML_PDF(SAPbouiCOM.Application sboappElectronicReception, SAPbobsCOM.Company oCompanyElectronicReception)
        {
            Funciones.Comunes DLLFunciones = new Funciones.Comunes();

            try
            {
                SAPbouiCOM.Form oForm = null;
                bool ExistForm = false;

                for (int i = 0; i < sboappElectronicReception.Forms.Count; i++)
                {
                    if (sboappElectronicReception.Forms.Item(i).UniqueID == "BOTVDR")
                    {
                        oForm = sboappElectronicReception.Forms.Item("BOTVDR");
                        ExistForm = true;
                    }
                }


                if (ExistForm)
                {
                    SAPbobsCOM.Recordset oRsProveedorTecnologico = (SAPbobsCOM.Recordset)oCompanyElectronicReception.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    string sGetProveedorTecnologico = null;

                    sGetProveedorTecnologico = DLLFunciones.GetStringXMLDocument(oCompanyElectronicReception, "eBilling", "eBilling", "GetProveedorTecnologico");
                    oRsProveedorTecnologico.DoQuery(sGetProveedorTecnologico);

                    if (oRsProveedorTecnologico.RecordCount > 0)
                    {
                        if (oRsProveedorTecnologico.Fields.Item("ProveedorTecnologico").Value.ToString() == "TFHKA")
                        {
                            DescargaXML_PDF_TFHKA(sboappElectronicReception, oCompanyElectronicReception, oForm);
                            
                        }
                        else if (oRsProveedorTecnologico.Fields.Item("ProveedorTecnologico").Value.ToString() == "FBE")
                        {
                            DescargaXML_PDF_TFHKA(sboappElectronicReception, oCompanyElectronicReception, oForm);
                        }
                    }
                }
                else
                {

                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        public void OpenFormParametrosIniciales(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _oCompany, string sMotor)
        {
            try
            {
                SAPbouiCOM.Form frm = null;
                bool ExistForm = false;

                for (int i = 0; i < _sboapp.Forms.Count; i++)
                {
                    if (_sboapp.Forms.Item(i).UniqueID == "BO_eBillingP")
                    {
                        frm = _sboapp.Forms.Item("BO_eBillingP");
                        ExistForm = true;
                    }
                }

                if (ExistForm)
                {
                    frm.Select();
                }
                else
                {
                    string ArchivoSRF = "RecepcionElectronica_ParametrosIniciales.srf";
                    DllFunciones.LoadFromXML(_sboapp, "BOElectronicReception", ref ArchivoSRF);

                    frm = _sboapp.Forms.Item("BO_eBillingP");

                    UpdateElectronicReceptionParameters(_oCompany, _sboapp, frm, sMotor);

                }
            }
            catch (Exception)
            {

                throw;
            }

        }

        public void UpdateElectronicReceptionParameters(SAPbobsCOM.Company oCompany, SAPbouiCOM.Application _sboapp, SAPbouiCOM.Form oFormVDBO, string _sMotor)
        {
            Funciones.Comunes DLLFunciones = new Funciones.Comunes();

            try
            {
                #region Variables y objetos

                SAPbouiCOM.ComboBox _cboPT = (SAPbouiCOM.ComboBox)oFormVDBO.Items.Item("Item_34").Specific;
                SAPbouiCOM.ComboBox cboFRE = (SAPbouiCOM.ComboBox)oFormVDBO.Items.Item("cboFRE").Specific;
                

                SAPbouiCOM.Folder oFolder1 = (SAPbouiCOM.Folder)oFormVDBO.Items.Item("Folder1").Specific;
                SAPbouiCOM.Folder oFolderTFHKA = (SAPbouiCOM.Folder)oFormVDBO.Items.Item("Folder3").Specific;
                SAPbouiCOM.Folder oFolderFacture = (SAPbouiCOM.Folder)oFormVDBO.Items.Item("Item_49").Specific;

                SAPbouiCOM.PictureBox oLogoBO = (SAPbouiCOM.PictureBox)oFormVDBO.Items.Item("LogoBO").Specific;
                                
                SAPbobsCOM.Recordset oGetFormaRecepcion = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oRsProveedorTecnologico = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string sGetProveedorTecnologico = null;
                string sGetFormaRecepcion = null;

                #endregion

                #region Selecciona Folders Proveedor Tecnologico

                sGetProveedorTecnologico = DLLFunciones.GetStringXMLDocument(oCompany, "BOElectronicReception", "ElectronicReception", "GetProveedorTecnologico");
                oRsProveedorTecnologico.DoQuery(sGetProveedorTecnologico);

                if (oRsProveedorTecnologico.RecordCount > 0)
                {
                    if (oRsProveedorTecnologico.Fields.Item("ProveedorTecnologico").Value.ToString() == "TFHKA")
                    {
                        oFolderTFHKA.Item.Visible = true;
                        oFolderFacture.Item.Visible = false;
                    }
                    else if (oRsProveedorTecnologico.Fields.Item("ProveedorTecnologico").Value.ToString() == "FBE")
                    {
                        oFolderTFHKA.Item.Visible = false;
                        oFolderFacture.Item.Visible = true;
                    }
                    else
                    {
                        oFolderTFHKA.Item.Visible = false;
                        oFolderFacture.Item.Visible = false;
                    }
                }
                else
                {
                    oFolderTFHKA.Item.Visible = false;
                    oFolderFacture.Item.Visible = false;
                }

                #endregion

                #region Asignacion Logo BO

                oLogoBO.Picture = (Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Core\\Images\\LogoBO20x20.bmp");

                #endregion

                #region Valores validos Forma Recepción

                sGetFormaRecepcion = DLLFunciones.GetStringXMLDocument(oCompany, "BOElectronicReception", "ElectronicReception", "GetFormaRecepcion");

                oGetFormaRecepcion.DoQuery(sGetFormaRecepcion);
                oGetFormaRecepcion.MoveFirst();

                for (int K = 0; oGetFormaRecepcion.RecordCount - 1 >= K; K++)
                {
                    cboFRE.ValidValues.Add(oGetFormaRecepcion.Fields.Item(0).Value.ToString(), oGetFormaRecepcion.Fields.Item(1).Value.ToString());
                    oGetFormaRecepcion.MoveNext();
                }

                DLLFunciones.liberarObjetos(oGetFormaRecepcion);

                #endregion
                
                oFormVDBO.DataBrowser.BrowseBy = "txtCode";

                oFormVDBO.Visible = true;
                oFormVDBO.Refresh();

                oFolder1.Select();

            }
            catch (Exception e)
            {
                DLLFunciones.sendErrorMessage(_sboapp, e);

            }
        }

        public void DocumentSearchDIAN(SAPbouiCOM.Application sboappElectronicReception, SAPbobsCOM.Company oCompanyElectronicReception, ItemEvent pVal, string ColID)
        {
            Funciones.Comunes DLLFunciones = new Funciones.Comunes();

            try
            {
                SAPbouiCOM.Form oForm = null;
                bool ExistForm = false;

                for (int i = 0; i < sboappElectronicReception.Forms.Count; i++)
                {
                    if (sboappElectronicReception.Forms.Item(i).UniqueID == "BOTVDR")
                    {
                        oForm = sboappElectronicReception.Forms.Item("BOTVDR");
                        ExistForm = true;
                    }
                }


                if (ExistForm)
                {
                    SAPbobsCOM.Recordset oRsProveedorTecnologico = (SAPbobsCOM.Recordset)oCompanyElectronicReception.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    string sGetProveedorTecnologico = null;

                    sGetProveedorTecnologico = DLLFunciones.GetStringXMLDocument(oCompanyElectronicReception, "eBilling", "eBilling", "GetProveedorTecnologico");
                    oRsProveedorTecnologico.DoQuery(sGetProveedorTecnologico);

                    if (oRsProveedorTecnologico.RecordCount > 0)
                    {
                        if (oRsProveedorTecnologico.Fields.Item("ProveedorTecnologico").Value.ToString() == "TFHKA")
                        {

                            SAPbouiCOM.Matrix oMatrixOPCH = (Matrix)oForm.Items.Item("MtxOPCH").Specific;

                            string sCUFE = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_17").Cells.Item(pVal.Row).Specific)).Value;

                            if (string.IsNullOrEmpty(sCUFE))
                            {

                            }
                            else
                            {

                                string URLDIAN = "https://catalogo-vpfe.dian.gov.co/document/searchqr?documentkey=" + sCUFE;
                                Process.Start(URLDIAN);
                                URLDIAN = null;

                                
                            }

                        }
                        else if (oRsProveedorTecnologico.Fields.Item("ProveedorTecnologico").Value.ToString() == "FBE")
                        {
                            DescargaXML_PDF_TFHKA(sboappElectronicReception, oCompanyElectronicReception, oForm);
                        }
                    }
                }
                else
                {

                }
            }
            catch (Exception)
            {

                throw;
            }






















        }

        public void ExportToExcel(SAPbouiCOM.Application sboappElectronicReception, SAPbobsCOM.Company oCompanyElectronicReception)
        {
            Funciones.Comunes DLLFunciones = new Funciones.Comunes();

            try
            {
                SAPbouiCOM.Form oForm = null;
                bool ExistForm = false;

                for (int i = 0; i < sboappElectronicReception.Forms.Count; i++)
                {
                    if (sboappElectronicReception.Forms.Item(i).UniqueID == "BOTVDR")
                    {
                        oForm = sboappElectronicReception.Forms.Item("BOTVDR");
                        ExistForm = true;
                    }
                }


                if (ExistForm)
                {
                    SAPbobsCOM.Recordset oRsProveedorTecnologico = (SAPbobsCOM.Recordset)oCompanyElectronicReception.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    string sGetProveedorTecnologico = null;

                    sGetProveedorTecnologico = DLLFunciones.GetStringXMLDocument(oCompanyElectronicReception, "eBilling", "eBilling", "GetProveedorTecnologico");
                    oRsProveedorTecnologico.DoQuery(sGetProveedorTecnologico);

                    if (oRsProveedorTecnologico.RecordCount > 0)
                    {
                        if (oRsProveedorTecnologico.Fields.Item("ProveedorTecnologico").Value.ToString() == "TFHKA")
                        {
                            ExportToExcel_TFHKA(sboappElectronicReception, oCompanyElectronicReception, oForm);

                        }
                        else if (oRsProveedorTecnologico.Fields.Item("ProveedorTecnologico").Value.ToString() == "FBE")
                        {
                            ExportToExcel_TFHKA(sboappElectronicReception, oCompanyElectronicReception, oForm);
                        }
                    }
                }
                else
                {

                }
            }
            catch (Exception)
            {

                throw;
            }
        }
        
        #region Proveedor Tecnologico TFHKA

        public void MatrixOpenFile_TFHKA(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Form _oFormVDR, ItemEvent pVal, string _ColUID)
        {
            if (_ColUID == "Col_20")
            {
                #region Open XML

                SAPbouiCOM.Matrix oMatrixOPCH = (Matrix)_oFormVDR.Items.Item("MtxOPCH").Specific;

                string sPath = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_12").Cells.Item(pVal.Row).Specific)).Value;

                if (string.IsNullOrEmpty(sPath))
                {

                }
                else
                {
                    if (pVal.Row == 0)
                    {

                    }
                    else
                    {
                        sPath = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_12").Cells.Item(pVal.Row).Specific)).Value;

                        System.Diagnostics.Process.Start(sPath);
                    }
                }

                #endregion
            }

            if (_ColUID == "Col_21")
            {
                #region Open PDF

                SAPbouiCOM.Matrix oMatrixOPCH = (Matrix)_oFormVDR.Items.Item("MtxOPCH").Specific;

                string sPath = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_13").Cells.Item(pVal.Row).Specific)).Value;

                if (string.IsNullOrEmpty(sPath))
                {

                }
                else
                {
                    if (pVal.Row == 0)
                    {

                    }
                    else
                    {
                        sPath = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_13").Cells.Item(pVal.Row).Specific)).Value;

                        System.Diagnostics.Process.Start(sPath);
                    }
                }

                #endregion
            }
        }

        private void DescargaDocumentosTFHKA(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Form _oFormVisorRepcecion)
        {
            try
            {
                int iReprocesar = DllFunciones.sendMessageBoxY_N(_sboapp, MessageSystemAddOn("FER10013", _sboapp, null));

                if (iReprocesar == 1)
                {
                    #region Consulta URL

                    string sGetModo = null;
                    string sURLEmision = null;
                    string sTokenEmpresa = null;
                    string sTokenPassword = null;
                    string sModo = null;
                    string sProtocoloComunicacion = null;
                    string sConsecutivoTFHKA = null;
                    string sFormaRecepcion = null;

                    SAPbobsCOM.Recordset oConsultarGetModo = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    SAPbobsCOM.Recordset oSyncDocsRecep = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    sGetModo = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetModoandURL");

                    oConsultarGetModo.DoQuery(sGetModo);

                    sURLEmision = Convert.ToString(oConsultarGetModo.Fields.Item("URLTFHKA").Value.ToString()) + "/ReceptorWS.svc?wsdl";
                    sModo = Convert.ToString(oConsultarGetModo.Fields.Item("Modo").Value.ToString());
                    sProtocoloComunicacion = Convert.ToString(oConsultarGetModo.Fields.Item("ProtocoloComunicacion").Value.ToString());
                    sTokenEmpresa = Convert.ToString(oConsultarGetModo.Fields.Item("TokenEmpresa").Value.ToString());
                    sTokenPassword = Convert.ToString(oConsultarGetModo.Fields.Item("TokenPassword").Value.ToString());
                    sConsecutivoTFHKA = Convert.ToString(oConsultarGetModo.Fields.Item("Consecutivo").Value.ToString());
                    sFormaRecepcion = Convert.ToString(oConsultarGetModo.Fields.Item("FormaRecepcion").Value.ToString());

                    DllFunciones.liberarObjetos(oConsultarGetModo);

                    #endregion

                    #region Variables y Objetos

                    string sPathImages = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\BOElectronicReception\\Images\\";

                    #endregion

                    #region Instanciacion parametros TFHKA

                    if (sProtocoloComunicacion == "HTTP")
                    {
                        BasicHttpBinding port = new BasicHttpBinding();
                    }
                    else if (sProtocoloComunicacion == "HTTPS")
                    {
                        BasicHttpsBinding port = new BasicHttpsBinding();
                    }

                    port.MaxBufferPoolSize = Int32.MaxValue;
                    port.MaxBufferSize = Int32.MaxValue;
                    port.MaxReceivedMessageSize = Int32.MaxValue;
                    port.ReaderQuotas.MaxStringContentLength = Int32.MaxValue;
                    port.SendTimeout = TimeSpan.FromMinutes(2);
                    port.ReceiveTimeout = TimeSpan.FromMinutes(2);

                    if (sProtocoloComunicacion == "HTTPS")
                    {
                        port.Security.Mode = BasicHttpSecurityMode.Transport;
                    }

                    EndpointAddress endPointEmision = new EndpointAddress(sURLEmision); //URL Recepcion

                    Recepcion21WS.ReceptorWSClient serviceClienTFHKAReception;

                    serviceClienTFHKAReception = new Recepcion21WS.ReceptorWSClient(port, endPointEmision);

                    ReceptorReporteRequest ParametrosConsultaTodosDocumentosRecepcion = new ReceptorReporteRequest();
                    ReceptorReporteStatusRequest ParametrosConsultaDocumentosRecepcion = new ReceptorReporteStatusRequest();

                    #endregion

                    #region Parametros Generales Reporte TFHKA

                    ParametrosConsultaTodosDocumentosRecepcion.tokenEmpresa = sTokenEmpresa;
                    ParametrosConsultaTodosDocumentosRecepcion.tokenPassword = sTokenPassword;

                    #endregion

                    #region Parametros Generales Reporte TFHKA por Estatus

                    ParametrosConsultaDocumentosRecepcion.tokenEmpresa = sTokenEmpresa;
                    ParametrosConsultaDocumentosRecepcion.tokenPassword = sTokenPassword;

                    #endregion

                    #region Consulta todos los documentos

                    string[] arrFormaRecepcion = { "XML", "y PDF" };

                    DllFunciones.ProgressBar(_oCompany, _sboapp, 10, 1, MessageSystemAddOn("FER10014", _sboapp, arrFormaRecepcion));

                    SincronizacionTodosDocumentosWS_TFHKA(ParametrosConsultaTodosDocumentosRecepcion, serviceClienTFHKAReception, _oCompany, _sboapp, oSyncDocsRecep, sFormaRecepcion);

                    #endregion

                    #region Consultas documentos por estado

                    DllFunciones.ProgressBar(_oCompany, _sboapp, 10, 1, MessageSystemAddOn("FER10015", _sboapp, arrFormaRecepcion));
                    SincronizacionWS_TFHKA("00", "Cargado", ParametrosConsultaDocumentosRecepcion, serviceClienTFHKAReception, _oCompany, _sboapp, oSyncDocsRecep);

                    DllFunciones.ProgressBar(_oCompany, _sboapp, 10, 1, MessageSystemAddOn("FER10016", _sboapp, arrFormaRecepcion));
                    SincronizacionWS_TFHKA("01", "Entregado", ParametrosConsultaDocumentosRecepcion, serviceClienTFHKAReception, _oCompany, _sboapp, oSyncDocsRecep);

                    DllFunciones.ProgressBar(_oCompany, _sboapp, 10, 1, MessageSystemAddOn("FER10017", _sboapp, arrFormaRecepcion));
                    SincronizacionWS_TFHKA("02", "Aceptación expresa (DIAN)", ParametrosConsultaDocumentosRecepcion, serviceClienTFHKAReception, _oCompany, _sboapp, oSyncDocsRecep);

                    DllFunciones.ProgressBar(_oCompany, _sboapp, 10, 1, MessageSystemAddOn("FER10018", _sboapp, arrFormaRecepcion));
                    SincronizacionWS_TFHKA("04", "Reclamo (DIAN)", ParametrosConsultaDocumentosRecepcion, serviceClienTFHKAReception, _oCompany, _sboapp, oSyncDocsRecep);

                    DllFunciones.ProgressBar(_oCompany, _sboapp, 10, 1, MessageSystemAddOn("FER10019", _sboapp, arrFormaRecepcion));
                    SincronizacionWS_TFHKA("10", "Acuse de recibo (DIAN", ParametrosConsultaDocumentosRecepcion, serviceClienTFHKAReception, _oCompany, _sboapp, oSyncDocsRecep);

                    DllFunciones.ProgressBar(_oCompany, _sboapp, 10, 1, MessageSystemAddOn("FER10020", _sboapp, arrFormaRecepcion));
                    SincronizacionWS_TFHKA("11", "Rechazado (DIAN)", ParametrosConsultaDocumentosRecepcion, serviceClienTFHKAReception, _oCompany, _sboapp, oSyncDocsRecep);

                    DllFunciones.ProgressBar(_oCompany, _sboapp, 10, 1, MessageSystemAddOn("FER10021", _sboapp, arrFormaRecepcion));
                    SincronizacionWS_TFHKA("12", "Recibo del bien y/o prestación del servicio", ParametrosConsultaDocumentosRecepcion, serviceClienTFHKAReception, _oCompany, _sboapp, oSyncDocsRecep);

                    DllFunciones.ProgressBar(_oCompany, _sboapp, 10, 1, MessageSystemAddOn("FER10022", _sboapp, arrFormaRecepcion));
                    SincronizacionWS_TFHKA("13", "Precargado", ParametrosConsultaDocumentosRecepcion, serviceClienTFHKAReception, _oCompany, _sboapp, oSyncDocsRecep);

                    #endregion

                    #region Carga Infortmacion en la Matrix

                    #region Variables y Objetos

                    string sPath;
                    string sInvoices = null;
                    string sCreditMemo = null;
                    string sDebitMemo = null;

                    int CantidadRegistos = 0;

                    #endregion

                    sPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

                    #region Consulta de documentos facturas, notas debito y notas credito a mostrar en matrix

                    SAPbouiCOM.Matrix oMatrixInvoice = (Matrix)_oFormVisorRepcecion.Items.Item("MtxOPCH").Specific;
                    SAPbouiCOM.Matrix oMatrixCreditMemo = (Matrix)_oFormVisorRepcecion.Items.Item("MtxORPC").Specific;
                    SAPbouiCOM.Matrix oMatrixDebitMemo = (Matrix)_oFormVisorRepcecion.Items.Item("MtxOPCHD").Specific;
                    SAPbouiCOM.EditText otxtND = (SAPbouiCOM.EditText)_oFormVisorRepcecion.Items.Item("txtND").Specific;
                    SAPbouiCOM.ComboBox oEstado = (SAPbouiCOM.ComboBox)_oFormVisorRepcecion.Items.Item("cboStado").Specific;

                    SAPbobsCOM.Recordset oRecorsetInvoices = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    SAPbobsCOM.Recordset oRecorsetCreditMemo = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    SAPbobsCOM.Recordset oRecorsetDebitMemo = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    SAPbouiCOM.DataTable oTableInvoices = _oFormVisorRepcecion.DataSources.DataTables.Item("DT_Invoices");
                    SAPbouiCOM.DataTable oTableCreditMemo = _oFormVisorRepcecion.DataSources.DataTables.Item("DT_CreditMemo");
                    SAPbouiCOM.DataTable oTableDebitMemo = _oFormVisorRepcecion.DataSources.DataTables.Item("DT_DebitMemo");

                    sInvoices = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetInvoices");
                    sInvoices = sInvoices.Replace("%PathImages%", sPathImages).Replace("%FI%", "20220108").Replace("%FF%", "20301231");

                    sCreditMemo = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetCreditMemo");
                    sInvoices = sInvoices.Replace("%PathImages%", sPathImages).Replace("%FI%", "20220108").Replace("%FF%", "20301231");

                    sDebitMemo = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetDebitMemo");
                    sInvoices = sInvoices.Replace("%PathImages%", sPathImages).Replace("%FI%", "20220108").Replace("%FF%", "20301231");

                    if (oEstado.Value == "-" || oEstado.Value == "")
                    {
                        sInvoices = sInvoices.Replace("%Estado%", "");
                    }
                    else
                    {
                        sInvoices = sInvoices.Replace("%Estado%", "AND \"U_BOTSDC\" = '" + oEstado.Value + "' ");
                    }

                    if (otxtND.Value == "")
                    {
                        sInvoices = sInvoices.Replace("%NumeroDocumento%", "");
                    }
                    else
                    {
                        sInvoices = sInvoices.Replace("%NumeroDocumento%", "AND \"U_BOTNF\" LIKE '%" + otxtND.Value + "%' ");
                    }

                    oRecorsetInvoices.DoQuery(sInvoices);
                    oRecorsetCreditMemo.DoQuery(sCreditMemo);
                    oRecorsetDebitMemo.DoQuery(sDebitMemo);

                    oTableInvoices.ExecuteQuery(sInvoices);
                    oTableCreditMemo.ExecuteQuery(sCreditMemo);
                    oTableDebitMemo.ExecuteQuery(sDebitMemo);

                    #endregion

                    CantidadRegistos = oRecorsetInvoices.RecordCount + oRecorsetCreditMemo.RecordCount + oRecorsetDebitMemo.RecordCount;

                    if (CantidadRegistos != 0)
                    {
                        #region Carga datos Matrix Facturas

                        if (oRecorsetInvoices.RecordCount > 0)
                        {
                            oMatrixInvoice.Clear();

                            oMatrixInvoice.Columns.Item("#").DataBind.Bind("DT_Invoices", "#");
                            oMatrixInvoice.Columns.Item("Col_0").DataBind.Bind("DT_Invoices", "Estado");
                            oMatrixInvoice.Columns.Item("Col_11").DataBind.Bind("DT_Invoices", "Estado_SAP");
                            oMatrixInvoice.Columns.Item("Col_9").DataBind.Bind("DT_Invoices", "DocEntry");
                            oMatrixInvoice.Columns.Item("Col_1").DataBind.Bind("DT_Invoices", "Num_Fac_Pro");
                            oMatrixInvoice.Columns.Item("Col_16").DataBind.Bind("DT_Invoices", "Serie_Numeracion");
                            oMatrixInvoice.Columns.Item("Col_2").DataBind.Bind("DT_Invoices", "Codigo_cliente");
                            oMatrixInvoice.Columns.Item("Col_3").DataBind.Bind("DT_Invoices", "Nombre_del_Cliente");
                            oMatrixInvoice.Columns.Item("Col_4").DataBind.Bind("DT_Invoices", "Fecha_Documento");
                            oMatrixInvoice.Columns.Item("Col_5").DataBind.Bind("DT_Invoices", "Fecha_vencimiento");
                            oMatrixInvoice.Columns.Item("Col_24").DataBind.Bind("DT_Invoices", "Condicion_de_Pago");
                            oMatrixInvoice.Columns.Item("Col_6").DataBind.Bind("DT_Invoices", "Total_documento");
                            oMatrixInvoice.Columns.Item("Col_17").DataBind.Bind("DT_Invoices", "CUFE");
                            oMatrixInvoice.Columns.Item("Col_18").DataBind.Bind("DT_Invoices", "Fecha_emision");
                            oMatrixInvoice.Columns.Item("Col_10").DataBind.Bind("DT_Invoices", "Hora_emision");
                            oMatrixInvoice.Columns.Item("Col_8").DataBind.Bind("DT_Invoices", "Fecha_recepcion");
                            oMatrixInvoice.Columns.Item("Col_19").Visible = false;
                            oMatrixInvoice.Columns.Item("Col_23").Visible = false;
                            oMatrixInvoice.Columns.Item("Col_20").DataBind.Bind("DT_Invoices", "DescargaXML");
                            oMatrixInvoice.Columns.Item("Col_12").DataBind.Bind("DT_Invoices", "RutaXML");
                            oMatrixInvoice.Columns.Item("Col_12").Visible = false;
                            oMatrixInvoice.Columns.Item("Col_21").DataBind.Bind("DT_Invoices", "DescargaPDF");
                            oMatrixInvoice.Columns.Item("Col_13").DataBind.Bind("DT_Invoices", "RutaPDF");
                            oMatrixInvoice.Columns.Item("Col_13").Visible = false;
                            oMatrixInvoice.Columns.Item("Col_22").DataBind.Bind("DT_Invoices", "ImageAceptar");
                            oMatrixInvoice.Columns.Item("Col_7").DataBind.Bind("DT_Invoices", "ImageCancelar");
                            oMatrixInvoice.Columns.Item("Col_14").DataBind.Bind("DT_Invoices", "TipoIdentificacionEmisor");
                            oMatrixInvoice.Columns.Item("Col_14").Visible = false;

                            oMatrixInvoice.LoadFromDataSource();

                            oMatrixInvoice.AutoResizeColumns();

                        }

                        #endregion

                        #region Carga datos Matrix Notas credito

                        if (oRecorsetCreditMemo.RecordCount > 0)
                        {
                            oMatrixCreditMemo.Clear();

                            oMatrixCreditMemo.Columns.Item("#").DataBind.Bind("DT_Invoices", "#");
                            oMatrixCreditMemo.Columns.Item("Col_0").DataBind.Bind("DT_Invoices", "Estado");
                            oMatrixCreditMemo.Columns.Item("Col_9").DataBind.Bind("DT_Invoices", "DocEntry");
                            oMatrixCreditMemo.Columns.Item("Col_1").DataBind.Bind("DT_Invoices", "Num_Fac_Pro");
                            //oMatrixCreditMemo.Columns.Item("Col_25").DataBind.Bind("DT_Invoices", "Num_Fac_Preeli");
                            //oMatrixCreditMemo.Columns.Item("Col_26").DataBind.Bind("DT_Invoices", "Num_Fac_SAP");
                            oMatrixCreditMemo.Columns.Item("Col_16").DataBind.Bind("DT_Invoices", "Serie_Numeracion");
                            oMatrixCreditMemo.Columns.Item("Col_2").DataBind.Bind("DT_Invoices", "Codigo_cliente");
                            oMatrixCreditMemo.Columns.Item("Col_3").DataBind.Bind("DT_Invoices", "Nombre_del_Cliente");
                            oMatrixCreditMemo.Columns.Item("Col_4").DataBind.Bind("DT_Invoices", "Fecha_Documento");
                            oMatrixCreditMemo.Columns.Item("Col_5").DataBind.Bind("DT_Invoices", "Fecha_vencimiento");
                            oMatrixCreditMemo.Columns.Item("Col_24").DataBind.Bind("DT_Invoices", "Condicion_Pago");
                            oMatrixCreditMemo.Columns.Item("Col_6").DataBind.Bind("DT_Invoices", "Total_documento");
                            oMatrixCreditMemo.Columns.Item("Col_17").DataBind.Bind("DT_Invoices", "CUFE");
                            oMatrixCreditMemo.Columns.Item("Col_18").DataBind.Bind("DT_Invoices", "Fecha_emision");
                            oMatrixCreditMemo.Columns.Item("Col_10").DataBind.Bind("DT_Invoices", "Hora_emision");
                            oMatrixCreditMemo.Columns.Item("Col_8").DataBind.Bind("DT_Invoices", "Fecha_recepcion");

                            oMatrixCreditMemo.LoadFromDataSource();

                            oMatrixCreditMemo.AutoResizeColumns();

                        }
                        #endregion

                        #region Carga datos Matrix Notas Debito

                        if (oRecorsetDebitMemo.RecordCount > 0)
                        {
                            oMatrixDebitMemo.Clear();

                            oMatrixDebitMemo.Columns.Item("#").DataBind.Bind("DT_Invoices", "#");
                            oMatrixDebitMemo.Columns.Item("Col_0").DataBind.Bind("DT_Invoices", "Estado");
                            oMatrixDebitMemo.Columns.Item("Col_9").DataBind.Bind("DT_Invoices", "DocEntry");
                            oMatrixDebitMemo.Columns.Item("Col_1").DataBind.Bind("DT_Invoices", "Num_Fac_Pro");
                            //oMatrixDebitMemo.Columns.Item("Col_25").DataBind.Bind("DT_Invoices", "Num_Fac_Preeli");
                            //oMatrixDebitMemo.Columns.Item("Col_26").DataBind.Bind("DT_Invoices", "Num_Fac_SAP");
                            oMatrixDebitMemo.Columns.Item("Col_16").DataBind.Bind("DT_Invoices", "Serie_Numeracion");
                            oMatrixDebitMemo.Columns.Item("Col_2").DataBind.Bind("DT_Invoices", "Codigo_cliente");
                            oMatrixDebitMemo.Columns.Item("Col_3").DataBind.Bind("DT_Invoices", "Nombre_del_Cliente");
                            oMatrixDebitMemo.Columns.Item("Col_4").DataBind.Bind("DT_Invoices", "Fecha_Documento");
                            oMatrixDebitMemo.Columns.Item("Col_5").DataBind.Bind("DT_Invoices", "Fecha_vencimiento");
                            oMatrixDebitMemo.Columns.Item("Col_24").DataBind.Bind("DT_Invoices", "Condicion_Pago");
                            oMatrixDebitMemo.Columns.Item("Col_6").DataBind.Bind("DT_Invoices", "Total_documento");
                            oMatrixDebitMemo.Columns.Item("Col_17").DataBind.Bind("DT_Invoices", "CUFE");
                            oMatrixDebitMemo.Columns.Item("Col_18").DataBind.Bind("DT_Invoices", "Fecha_emision");
                            oMatrixDebitMemo.Columns.Item("Col_10").DataBind.Bind("DT_Invoices", "Hora_emision");
                            oMatrixDebitMemo.Columns.Item("Col_8").DataBind.Bind("DT_Invoices", "Fecha_recepcion");

                            oMatrixDebitMemo.LoadFromDataSource();

                            oMatrixDebitMemo.AutoResizeColumns();

                        }
                        #endregion

                    }
                    else
                    {
                        DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10023", _sboapp, arrFormaRecepcion));
                    }

                    DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, MessageSystemAddOn("FER10024", _sboapp, arrFormaRecepcion));

                    #endregion

                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        public void SincronizacionTodosDocumentosWS_TFHKA(ReceptorReporteRequest ParametrosConsultaTodosDocumentosRecepcion, Recepcion21WS.ReceptorWSClient _serviceClienTFHKAReception, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Application sboapp, SAPbobsCOM.Recordset oSyncDocsRecep, string sFormaRecepcion )
        {
            string sSyncDocsRecepOriginal = null;
            string sSyncDocsRecepCopia = null;
            string sGetQuantityDocumentsReceipt = null;
            string sGetModo = null;
            string sRutaXML = null;
            string sRutaPDF = null;

            var ResponsiveDocumentsReceptor = _serviceClienTFHKAReception.Reporte(ParametrosConsultaTodosDocumentosRecepcion);

            if (ResponsiveDocumentsReceptor.codigo == 200)
            {
                #region Sincronizando documentos con el proveedor tecnologico                
                
                #region Queries de Consulta y actualizacion
                
                sSyncDocsRecepOriginal = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "SyncDocsRecep");

                sGetQuantityDocumentsReceipt = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetQuantityDocumentsReceipt");

                SAPbobsCOM.Recordset oGetQuantityDocumentsReceipt = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                oGetQuantityDocumentsReceipt.DoQuery(sGetQuantityDocumentsReceipt);

                #endregion

                #region Queries consulta parametros

                SAPbobsCOM.Recordset oConsultarGetModo = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                sGetModo = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetModoandURL");

                sGetModo = sGetModo.Replace("%Estado%", "\"U_BO_Status\" = 'Y'").Replace("%DocEntry%", " ");

                oConsultarGetModo.DoQuery(sGetModo);

                sRutaXML = Convert.ToString(oConsultarGetModo.Fields.Item("RutaXML").Value.ToString());
                sRutaPDF = Convert.ToString(oConsultarGetModo.Fields.Item("RutaPDF").Value.ToString());

                #endregion

                string[] arrFormaRecepcion = { "XML", "y PDF" };

                DllFunciones.ProgressBar(_oCompany, sboapp, 10, 1, MessageSystemAddOn("FER10025", sboapp, arrFormaRecepcion));

                if (oGetQuantityDocumentsReceipt.RecordCount > 0)
                {
                    #region Recibe los ultimos 500 registros

                    for (int i = 0; i < ResponsiveDocumentsReceptor.documentoselectronicos.Count(); i++)
                    {
                        #region Inserta el documentos en la base de datos 

                        sSyncDocsRecepCopia = null;

                        sSyncDocsRecepCopia = sSyncDocsRecepOriginal;

                        sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%ProveedorTechnologico%", "TFHKA").Replace("%ConsecutivoTFHKA%", ResponsiveDocumentsReceptor.documentoselectronicos[i].correlativoempresa.ToString()).Replace("%CUFE%", ResponsiveDocumentsReceptor.documentoselectronicos[i].cufe.ToString());

                        #region StatusDIANCodigo

                        if (string.IsNullOrEmpty(ResponsiveDocumentsReceptor.documentoselectronicos[i].estatusDIANcodigo.ToString()))
                        {
                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%StatusDIANCodigo%'", "NULL");
                        }
                        else
                        {
                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%StatusDIANCodigo%", ResponsiveDocumentsReceptor.documentoselectronicos[i].estatusDIANcodigo.ToString());
                        }

                        #endregion

                        #region StatusDIANDescripcion

                        if (string.IsNullOrEmpty(ResponsiveDocumentsReceptor.documentoselectronicos[i].estatusDIANdescripcion.ToString()))
                        {
                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%StatusDIANDescripcion%'", "NULL");
                        }
                        else
                        {
                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%StatusDIANDescripcion%", ResponsiveDocumentsReceptor.documentoselectronicos[i].estatusDIANdescripcion.ToString());
                        }

                        #endregion

                        #region StatusDIANFecha

                        if (string.IsNullOrEmpty(ResponsiveDocumentsReceptor.documentoselectronicos[i].estatusDIANfecha.ToString()))
                        {
                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%StatusDIANFecha%'", "NULL");
                        }
                        else
                        {
                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%StatusDIANFecha%", ResponsiveDocumentsReceptor.documentoselectronicos[i].estatusDIANfecha.ToString());
                        }

                        #endregion

                        #region FechaEmision

                        if (string.IsNullOrEmpty(ResponsiveDocumentsReceptor.documentoselectronicos[i].fechaemision.ToString()))
                        {
                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%FechaEmision%'", "NULL");
                        }
                        else
                        {
                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%FechaEmision%", ResponsiveDocumentsReceptor.documentoselectronicos[i].fechaemision.ToString());
                        }

                        #endregion

                        #region FechaRecepcion

                        if (string.IsNullOrEmpty(ResponsiveDocumentsReceptor.documentoselectronicos[i].fecharecepcion.ToString()))
                        {
                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%FechaRecepcion%'", "NULL");
                        }
                        else
                        {
                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%FechaRecepcion%", ResponsiveDocumentsReceptor.documentoselectronicos[i].fecharecepcion.ToString());
                        }

                        #endregion

                        #region HoraEmision

                        if (string.IsNullOrEmpty(ResponsiveDocumentsReceptor.documentoselectronicos[i].horaemision.ToString()))
                        {
                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%HoraEmision%'", "NULL");
                        }
                        else
                        {
                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%HoraEmision%", ResponsiveDocumentsReceptor.documentoselectronicos[i].horaemision.ToString());
                        }

                        #endregion

                        #region MontoTotal

                        if (string.IsNullOrEmpty(ResponsiveDocumentsReceptor.documentoselectronicos[i].montototal.ToString()))
                        {
                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%MontoTotal%'", "NULL");
                        }
                        else
                        {
                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%MontoTotal%", ResponsiveDocumentsReceptor.documentoselectronicos[i].montototal.ToString());
                        }

                        #endregion

                        #region NumeroFactura

                        if (string.IsNullOrEmpty(ResponsiveDocumentsReceptor.documentoselectronicos[i].numerodocumento.ToString()))
                        {
                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%NumeroFactura%'", "NULL");
                        }
                        else
                        {
                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%NumeroFactura%", ResponsiveDocumentsReceptor.documentoselectronicos[i].numerodocumento.ToString());
                        }

                        #endregion

                        #region NumeroIdentificacion

                        if (string.IsNullOrEmpty(ResponsiveDocumentsReceptor.documentoselectronicos[i].numeroidentificacion.ToString()))
                        {
                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%NumeroIdentificacion%'", "NULL");
                        }
                        else
                        {
                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%NumeroIdentificacion%", ResponsiveDocumentsReceptor.documentoselectronicos[i].numeroidentificacion.ToString());
                        }

                        #endregion

                        #region RazonSocial

                        if (string.IsNullOrEmpty(ResponsiveDocumentsReceptor.documentoselectronicos[i].razonsocial.ToString()))
                        {
                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%RazonSocial%'", "NULL");
                        }
                        else
                        {
                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%RazonSocial%", ResponsiveDocumentsReceptor.documentoselectronicos[i].razonsocial.ToString());
                        }

                        #endregion

                        #region tipodocumento

                        if (string.IsNullOrEmpty(ResponsiveDocumentsReceptor.documentoselectronicos[i].tipodocumento.ToString()))
                        {
                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%TipoDocumento%'", "NULL");
                        }
                        else
                        {
                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%TipoDocumento%", ResponsiveDocumentsReceptor.documentoselectronicos[i].tipodocumento.ToString());
                        }

                        #endregion

                        #region TipoEmisor

                        if (string.IsNullOrEmpty(ResponsiveDocumentsReceptor.documentoselectronicos[i].tipoemisor.ToString()))
                        {
                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%TipoEmisor%'", "NULL");
                        }
                        else
                        {
                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%TipoEmisor%", ResponsiveDocumentsReceptor.documentoselectronicos[i].tipoemisor.ToString());
                        }

                        #endregion

                        #region TipoIdentidad

                        if (string.IsNullOrEmpty(ResponsiveDocumentsReceptor.documentoselectronicos[i].tipoidentidad.ToString()))
                        {
                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%TipoIdentidad%'", "NULL");
                        }
                        else
                        {
                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%TipoIdentidad%", ResponsiveDocumentsReceptor.documentoselectronicos[i].tipoidentidad.ToString());
                        }

                        #endregion

                        #region Codigo Estatus DIAN

                        sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%prmCodEstatusDIAN%'", "NULL");

                        #endregion                       

                        #region Descarga XML y PDF

                        if (sFormaRecepcion == "B")
                        {
                            DescargaXML_PDF_WS_TFHKA(sboapp, _oCompany, sFormaRecepcion, ResponsiveDocumentsReceptor.documentoselectronicos[i].numeroidentificacion.ToString(), ResponsiveDocumentsReceptor.documentoselectronicos[i].numerodocumento.ToString(), ResponsiveDocumentsReceptor.documentoselectronicos[i].tipoidentidad.ToString());

                            string sRutaArchivoXML = sRutaXML + "\\" + ResponsiveDocumentsReceptor.documentoselectronicos[i].numeroidentificacion.ToString() + "_" + ResponsiveDocumentsReceptor.documentoselectronicos[i].numerodocumento.ToString() + ".xml";
                            
                            if (File.Exists(sRutaArchivoXML))
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%PathXML%", sRutaArchivoXML);
                            }
                            else
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%PathXML%'", sRutaArchivoXML);
                            }

                        }
                        else if (sFormaRecepcion == "C")
                        {
                            DescargaXML_PDF_WS_TFHKA(sboapp, _oCompany, sFormaRecepcion, ResponsiveDocumentsReceptor.documentoselectronicos[i].numeroidentificacion.ToString(), ResponsiveDocumentsReceptor.documentoselectronicos[i].numerodocumento.ToString(), ResponsiveDocumentsReceptor.documentoselectronicos[i].tipoidentidad.ToString());

                            string sRutaArchivoXML = sRutaXML + "\\" + ResponsiveDocumentsReceptor.documentoselectronicos[i].numeroidentificacion.ToString() + "_" + ResponsiveDocumentsReceptor.documentoselectronicos[i].numerodocumento.ToString() + ".xml";
                            string sRutaArchivoPDF = sRutaPDF + "\\" + ResponsiveDocumentsReceptor.documentoselectronicos[i].numeroidentificacion.ToString() + "_" + ResponsiveDocumentsReceptor.documentoselectronicos[i].numerodocumento.ToString() + ".pdf";

                            if (File.Exists(sRutaArchivoXML))
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%PathXML%", sRutaArchivoXML);
                            }
                            else
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%PathXML%'", "NULL");
                            }

                            if (File.Exists(sRutaArchivoPDF))
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%PathPDF%", sRutaArchivoPDF);
                            }
                            else
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%PathPDF%'", "NULL");
                            }

                        }

                        #endregion

                        #region Obtiene Elemento "OrderReference"                        

                        var sResponsePO = GetElementXML_DIAN(sboapp, _oCompany, "OrderReference", sRutaXML, ResponsiveDocumentsReceptor.documentoselectronicos[i].numeroidentificacion.ToString(), ResponsiveDocumentsReceptor.documentoselectronicos[i].numerodocumento.ToString());

                        if (string.IsNullOrEmpty(sResponsePO))
                        {
                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%PO%'", "NULL");
                        }
                        else
                        {
                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%PO%", sResponsePO.ToString());
                        }

                        #endregion

                        #region Obtiene Elemento "PaymentMeans_ID"                        

                        var sResponsePaymentMeans_ID = GetElementXML_DIAN(sboapp, _oCompany, "PaymentMeans_ID", sRutaXML, ResponsiveDocumentsReceptor.documentoselectronicos[i].numeroidentificacion.ToString(), ResponsiveDocumentsReceptor.documentoselectronicos[i].numerodocumento.ToString());

                        if (string.IsNullOrEmpty(sResponsePaymentMeans_ID))
                        {
                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%MedioPago%'", "NULL");
                        }
                        else
                        {
                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%MedioPago%", sResponsePaymentMeans_ID.ToString());
                        }

                        #endregion

                        oSyncDocsRecep.DoQuery(sSyncDocsRecepCopia);

                        #endregion
                    }
                    #endregion
                }
                else
                {
                    #region Inserta todos lo registros desde el inicio de la empresa

                    int iConsecutivoUltimo = Convert.ToInt32(ResponsiveDocumentsReceptor.ultimoEnviado.ToString());
                    bool PrimeraEjecucion = true;

                    do
                    {
                        if (PrimeraEjecucion)
                        {

                        }
                        else
                        {
                            ParametrosConsultaTodosDocumentosRecepcion.consecutivo = iConsecutivoUltimo.ToString();
                            ResponsiveDocumentsReceptor = _serviceClienTFHKAReception.Reporte(ParametrosConsultaTodosDocumentosRecepcion);
                        }

                        for (int i = 0; i < ResponsiveDocumentsReceptor.documentoselectronicos.Count(); i++)
                        {
                            #region Inserta el documentos en la base de datos 

                            sSyncDocsRecepCopia = null;

                            sSyncDocsRecepCopia = sSyncDocsRecepOriginal;

                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%ProveedorTechnologico%", "TFHKA").Replace("%ConsecutivoTFHKA%", ResponsiveDocumentsReceptor.documentoselectronicos[i].correlativoempresa.ToString()).Replace("%CUFE%", ResponsiveDocumentsReceptor.documentoselectronicos[i].cufe.ToString()).Replace("'%PathXML%'", "NULL").Replace("'%PathPDF%'", "NULL");

                            #region StatusDIANCodigo

                            if (string.IsNullOrEmpty(ResponsiveDocumentsReceptor.documentoselectronicos[i].estatusDIANcodigo.ToString()))
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%StatusDIANCodigo%'", "NULL");
                            }
                            else
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%StatusDIANCodigo%", ResponsiveDocumentsReceptor.documentoselectronicos[i].estatusDIANcodigo.ToString());
                            }

                            #endregion

                            #region StatusDIANDescripcion

                            if (string.IsNullOrEmpty(ResponsiveDocumentsReceptor.documentoselectronicos[i].estatusDIANdescripcion.ToString()))
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%StatusDIANDescripcion%'", "NULL");
                            }
                            else
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%StatusDIANDescripcion%", ResponsiveDocumentsReceptor.documentoselectronicos[i].estatusDIANdescripcion.ToString());
                            }

                            #endregion

                            #region StatusDIANFecha

                            if (string.IsNullOrEmpty(ResponsiveDocumentsReceptor.documentoselectronicos[i].estatusDIANfecha.ToString()))
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%StatusDIANFecha%'", "NULL");
                            }
                            else
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%StatusDIANFecha%", ResponsiveDocumentsReceptor.documentoselectronicos[i].estatusDIANfecha.ToString());
                            }

                            #endregion

                            #region FechaEmision

                            if (string.IsNullOrEmpty(ResponsiveDocumentsReceptor.documentoselectronicos[i].fechaemision.ToString()))
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%FechaEmision%'", "NULL");
                            }
                            else
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%FechaEmision%", ResponsiveDocumentsReceptor.documentoselectronicos[i].fechaemision.ToString());
                            }

                            #endregion

                            #region FechaRecepcion

                            if (string.IsNullOrEmpty(ResponsiveDocumentsReceptor.documentoselectronicos[i].fecharecepcion.ToString()))
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%FechaRecepcion%'", "NULL");
                            }
                            else
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%FechaRecepcion%", ResponsiveDocumentsReceptor.documentoselectronicos[i].fecharecepcion.ToString());
                            }

                            #endregion

                            #region HoraEmision

                            if (string.IsNullOrEmpty(ResponsiveDocumentsReceptor.documentoselectronicos[i].horaemision.ToString()))
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%HoraEmision%'", "NULL");
                            }
                            else
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%HoraEmision%", ResponsiveDocumentsReceptor.documentoselectronicos[i].horaemision.ToString());
                            }

                            #endregion

                            #region MontoTotal

                            if (string.IsNullOrEmpty(ResponsiveDocumentsReceptor.documentoselectronicos[i].montototal.ToString()))
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%MontoTotal%'", "NULL");
                            }
                            else
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%MontoTotal%", ResponsiveDocumentsReceptor.documentoselectronicos[i].montototal.ToString());
                            }

                            #endregion

                            #region NumeroFactura

                            if (string.IsNullOrEmpty(ResponsiveDocumentsReceptor.documentoselectronicos[i].numerodocumento.ToString()))
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%NumeroFactura%'", "NULL");
                            }
                            else
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%NumeroFactura%", ResponsiveDocumentsReceptor.documentoselectronicos[i].numerodocumento.ToString());
                            }

                            #endregion

                            #region NumeroIdentificacion

                            if (string.IsNullOrEmpty(ResponsiveDocumentsReceptor.documentoselectronicos[i].numeroidentificacion.ToString()))
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%NumeroIdentificacion%'", "NULL");
                            }
                            else
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%NumeroIdentificacion%", ResponsiveDocumentsReceptor.documentoselectronicos[i].numeroidentificacion.ToString());
                            }

                            #endregion

                            #region RazonSocial

                            if (string.IsNullOrEmpty(ResponsiveDocumentsReceptor.documentoselectronicos[i].razonsocial.ToString()))
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%RazonSocial%'", "NULL");
                            }
                            else
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%RazonSocial%", ResponsiveDocumentsReceptor.documentoselectronicos[i].razonsocial.ToString());
                            }

                            #endregion

                            #region tipodocumento

                            if (string.IsNullOrEmpty(ResponsiveDocumentsReceptor.documentoselectronicos[i].tipodocumento.ToString()))
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%TipoDocumento%'", "NULL");
                            }
                            else
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%TipoDocumento%", ResponsiveDocumentsReceptor.documentoselectronicos[i].tipodocumento.ToString());
                            }

                            #endregion

                            #region TipoEmisor

                            if (string.IsNullOrEmpty(ResponsiveDocumentsReceptor.documentoselectronicos[i].tipoemisor.ToString()))
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%TipoEmisor%'", "NULL");
                            }
                            else
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%TipoEmisor%", ResponsiveDocumentsReceptor.documentoselectronicos[i].tipoemisor.ToString());
                            }

                            #endregion

                            #region TipoIdentidad

                            if (string.IsNullOrEmpty(ResponsiveDocumentsReceptor.documentoselectronicos[i].tipoidentidad.ToString()))
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%TipoIdentidad%'", "NULL");
                            }
                            else
                            {
                                sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%TipoIdentidad%", ResponsiveDocumentsReceptor.documentoselectronicos[i].tipoidentidad.ToString());
                            }

                            #endregion

                            #region Codigo Estatus DIAN

                            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%prmCodEstatusDIAN%'", "NULL");

                            #endregion

                            oSyncDocsRecep.DoQuery(sSyncDocsRecepCopia);

                            if (sFormaRecepcion == "B")
                            {
                                DescargaXML_PDF_WS_TFHKA(sboapp, _oCompany, sFormaRecepcion, ResponsiveDocumentsReceptor.documentoselectronicos[i].numeroidentificacion.ToString(), ResponsiveDocumentsReceptor.documentoselectronicos[i].numerodocumento.ToString(), ResponsiveDocumentsReceptor.documentoselectronicos[i].tipoidentidad.ToString());
                            }
                            else if (sFormaRecepcion == "C")
                            {
                                DescargaXML_PDF_WS_TFHKA(sboapp, _oCompany, sFormaRecepcion, ResponsiveDocumentsReceptor.documentoselectronicos[i].numeroidentificacion.ToString(), ResponsiveDocumentsReceptor.documentoselectronicos[i].numerodocumento.ToString(), ResponsiveDocumentsReceptor.documentoselectronicos[i].tipoidentidad.ToString());
                            }

                            #endregion
                        }

                        iConsecutivoUltimo = Convert.ToInt32(ResponsiveDocumentsReceptor.ultimoEnviado.ToString());

                        PrimeraEjecucion = false;

                    } while (iConsecutivoUltimo > 1);

                    #endregion
                }

                DllFunciones.liberarObjetos(oGetQuantityDocumentsReceipt);

                #endregion
            }

        }

        public void SincronizacionWS_TFHKA(string _sCodigoStatusDIAN, string _DescripcionEstatusDIAN, ReceptorReporteStatusRequest ParametrosConsultaDocumentosRecepcion, Recepcion21WS.ReceptorWSClient _serviceClienTFHKAReception, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Application sboapp, SAPbobsCOM.Recordset oSyncDocsRecep)
        {
            string sCodigoStatusDIAN = _sCodigoStatusDIAN;
            string sSyncDocsRecepOriginal = null;
            string sSyncDocsRecepCopia = null;

            ParametrosConsultaDocumentosRecepcion.status_code = sCodigoStatusDIAN;

            var ResponsiveReportReception00 = _serviceClienTFHKAReception.ReporteStatus(ParametrosConsultaDocumentosRecepcion);

            if (ResponsiveReportReception00.codigo == 200)
            {
                #region Sincronizando documentos con el proveedor tecnologico                

                sSyncDocsRecepOriginal = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "SyncDocsRecep");

                for (int i = 0; i < ResponsiveReportReception00.documentoselectronicos.Count(); i++)
                {

                    sSyncDocsRecepCopia = null;

                    sSyncDocsRecepCopia = sSyncDocsRecepOriginal;

                    sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%ProveedorTechnologico%", "TFHKA").Replace("%ConsecutivoTFHKA%", ResponsiveReportReception00.documentoselectronicos[i].correlativoempresa.ToString()).Replace("%CUFE%", ResponsiveReportReception00.documentoselectronicos[i].cufe.ToString());

                    #region StatusDIANCodigo

                    if (string.IsNullOrEmpty(ResponsiveReportReception00.documentoselectronicos[i].estatusDIANcodigo.ToString()))
                    {
                        sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%StatusDIANCodigo%'", _sCodigoStatusDIAN);
                    }
                    else
                    {
                        sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%StatusDIANCodigo%", ResponsiveReportReception00.documentoselectronicos[i].estatusDIANcodigo.ToString());
                    }

                    #endregion

                    #region StatusDIANDescripcion

                    if (string.IsNullOrEmpty(ResponsiveReportReception00.documentoselectronicos[i].estatusDIANdescripcion.ToString()))
                    {
                        sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%StatusDIANDescripcion%'", "NULL");
                    }
                    else
                    {
                        sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%StatusDIANDescripcion%", ResponsiveReportReception00.documentoselectronicos[i].estatusDIANdescripcion.ToString());
                    }

                    #endregion

                    #region StatusDIANFecha

                    if (string.IsNullOrEmpty(ResponsiveReportReception00.documentoselectronicos[i].estatusDIANfecha.ToString()))
                    {
                        sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%StatusDIANFecha%'", "NULL");
                    }
                    else
                    {
                        sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%StatusDIANFecha%", ResponsiveReportReception00.documentoselectronicos[i].estatusDIANfecha.ToString());
                    }

                    #endregion

                    #region FechaEmision

                    if (string.IsNullOrEmpty(ResponsiveReportReception00.documentoselectronicos[i].fechaemision.ToString()))
                    {
                        sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%FechaEmision%'", "NULL");
                    }
                    else
                    {
                        sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%FechaEmision%", ResponsiveReportReception00.documentoselectronicos[i].fechaemision.ToString());
                    }

                    #endregion

                    #region FechaRecepcion

                    if (string.IsNullOrEmpty(ResponsiveReportReception00.documentoselectronicos[i].fecharecepcion.ToString()))
                    {
                        sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%FechaRecepcion%'", "NULL");
                    }
                    else
                    {
                        sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%FechaRecepcion%", ResponsiveReportReception00.documentoselectronicos[i].fecharecepcion.ToString());
                    }

                    #endregion

                    #region HoraEmision

                    if (string.IsNullOrEmpty(ResponsiveReportReception00.documentoselectronicos[i].horaemision.ToString()))
                    {
                        sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%HoraEmision%'", "NULL");
                    }
                    else
                    {
                        sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%HoraEmision%", ResponsiveReportReception00.documentoselectronicos[i].horaemision.ToString());
                    }

                    #endregion

                    #region MontoTotal

                    if (string.IsNullOrEmpty(ResponsiveReportReception00.documentoselectronicos[i].montototal.ToString()))
                    {
                        sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%MontoTotal%'", "NULL");
                    }
                    else
                    {
                        sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%MontoTotal%", ResponsiveReportReception00.documentoselectronicos[i].montototal.ToString());
                    }

                    #endregion

                    #region NumeroFactura

                    if (string.IsNullOrEmpty(ResponsiveReportReception00.documentoselectronicos[i].numerodocumento.ToString()))
                    {
                        sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%NumeroFactura%'", "NULL");
                    }
                    else
                    {
                        sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%NumeroFactura%", ResponsiveReportReception00.documentoselectronicos[i].numerodocumento.ToString());
                    }

                    #endregion

                    #region NumeroIdentificacion

                    if (string.IsNullOrEmpty(ResponsiveReportReception00.documentoselectronicos[i].numeroidentificacion.ToString()))
                    {
                        sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%NumeroIdentificacion%'", "NULL");
                    }
                    else
                    {
                        sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%NumeroIdentificacion%", ResponsiveReportReception00.documentoselectronicos[i].numeroidentificacion.ToString());
                    }

                    #endregion

                    #region RazonSocial

                    if (string.IsNullOrEmpty(ResponsiveReportReception00.documentoselectronicos[i].razonsocial.ToString()))
                    {
                        sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%RazonSocial%'", "NULL");
                    }
                    else
                    {
                        sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%RazonSocial%", ResponsiveReportReception00.documentoselectronicos[i].razonsocial.ToString());
                    }

                    #endregion

                    #region tipodocumento

                    if (string.IsNullOrEmpty(ResponsiveReportReception00.documentoselectronicos[i].tipodocumento.ToString()))
                    {
                        sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%TipoDocumento%'", "NULL");
                    }
                    else
                    {
                        sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%TipoDocumento%", ResponsiveReportReception00.documentoselectronicos[i].tipodocumento.ToString());
                    }

                    #endregion

                    #region TipoEmisor

                    if (string.IsNullOrEmpty(ResponsiveReportReception00.documentoselectronicos[i].tipoemisor.ToString()))
                    {
                        sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%TipoEmisor%'", "NULL");
                    }
                    else
                    {
                        sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%TipoEmisor%", ResponsiveReportReception00.documentoselectronicos[i].tipoemisor.ToString());
                    }

                    #endregion

                    #region TipoIdentidad

                    if (string.IsNullOrEmpty(ResponsiveReportReception00.documentoselectronicos[i].tipoidentidad.ToString()))
                    {
                        sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%TipoIdentidad%'", "NULL");
                    }
                    else
                    {
                        sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%TipoIdentidad%", ResponsiveReportReception00.documentoselectronicos[i].tipoidentidad.ToString());
                    }

                    #endregion

                    #region Codigo Estatus DIAN

                    sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%prmCodEstatusDIAN%", sCodigoStatusDIAN);

                    #endregion

                    #region Order Reference

                    sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%PO%'", "NULL");

                    #endregion

                    #region Medio Pago

                    sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%MedioPago%'", "NULL");

                    #endregion

                    oSyncDocsRecep.DoQuery(sSyncDocsRecepCopia);

                }

                #endregion
            }

        }

        private void DescargaXML_PDF_WS_TFHKA(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _oCompany, string sFormaRecepcion, string U_BOTNI, string U_BOTNF, string U_BOTTI)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {
                #region Consulta URL

                string sGetModo = null;
                string sURLRecepcion = null;
                string sModo = null;
                string sRutaXML = null;
                string sRutaPDF = null;
                string sProtocoloComunicacion = null;
                string sTokenEmpresa = null;
                string sTokenPassword = null;                
                       
                string sSyncAttachmentOrigin = null;

                SAPbobsCOM.Recordset oConsultarGetModo = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);                

                sGetModo = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetModoandURL");

                sGetModo = sGetModo.Replace("%Estado%", "\"U_BO_Status\" = 'Y'").Replace("%DocEntry%", " ");

                oConsultarGetModo.DoQuery(sGetModo);

                sURLRecepcion = Convert.ToString(oConsultarGetModo.Fields.Item("URLTFHKA").Value.ToString()) + "/ReceptorWS.svc?wsdl";
                sModo = Convert.ToString(oConsultarGetModo.Fields.Item("Modo").Value.ToString());

                sProtocoloComunicacion = Convert.ToString(oConsultarGetModo.Fields.Item("ProtocoloComunicacion").Value.ToString());
                sTokenEmpresa = Convert.ToString(oConsultarGetModo.Fields.Item("TokenEmpresa").Value.ToString());
                sTokenPassword = Convert.ToString(oConsultarGetModo.Fields.Item("TokenPassword").Value.ToString());
                sRutaXML = Convert.ToString(oConsultarGetModo.Fields.Item("RutaXML").Value.ToString());
                sRutaPDF = Convert.ToString(oConsultarGetModo.Fields.Item("RutaPDF").Value.ToString());

                DllFunciones.liberarObjetos(oConsultarGetModo);

                #endregion

                string sRutaArchivoXML = sRutaXML + "\\" + U_BOTNI + "_" + U_BOTNF + ".xml";
                string sRutaArchivoPDF = sRutaPDF + "\\" + U_BOTNI + "_" + U_BOTNF + ".pdf";

                if (File.Exists(sRutaArchivoXML) && File.Exists(sRutaArchivoPDF))
                {
                    
                }
                else
                {
                    #region Instanciacion parametros TFHKA

                    if (sProtocoloComunicacion == "HTTP")
                    {
                        BasicHttpBinding port = new BasicHttpBinding();
                    }
                    else if (sProtocoloComunicacion == "HTTPS")
                    {
                        BasicHttpsBinding port = new BasicHttpsBinding();
                    }

                    port.MaxBufferPoolSize = Int32.MaxValue;
                    port.MaxBufferSize = Int32.MaxValue;
                    port.MaxReceivedMessageSize = Int32.MaxValue;
                    port.ReaderQuotas.MaxStringContentLength = Int32.MaxValue;
                    port.SendTimeout = TimeSpan.FromMinutes(2);
                    port.ReceiveTimeout = TimeSpan.FromMinutes(2);

                    if (sProtocoloComunicacion == "HTTPS")
                    {
                        port.Security.Mode = BasicHttpSecurityMode.Transport;
                    }

                    EndpointAddress endPointEmision = new EndpointAddress(sURLRecepcion); //URL

                    Recepcion21WS.ReceptorWSClient serviceClienTFHKAReception;
                    serviceClienTFHKAReception = new Recepcion21WS.ReceptorWSClient(port, endPointEmision);

                    #endregion

                    #region Parametros generales Reporte

                    ReceptorRequestGeneral ParametrosDownload = new ReceptorRequestGeneral();

                    ParametrosDownload.tokenEmpresa = sTokenEmpresa;
                    ParametrosDownload.tokenPassword = sTokenPassword;

                    #endregion

                    #region Query Sincronizacion documentso

                    SAPbobsCOM.Recordset oRsSyncDocument = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    sSyncAttachmentOrigin = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "SyncAttachment");

                    #endregion

                    #region Descarga Documentos                     

                    Recepcion21WS.FileDownloadResponse DownloadXML;
                    Recepcion21WS.FileDownloadResponse DownloadPDF;

                    #region Parametros Complementarios Reporte

                    ParametrosDownload.identificadorEmisor = U_BOTNI;
                    ParametrosDownload.numeroDocumento = U_BOTNF;
                    ParametrosDownload.tipoIdentificacionemisor = U_BOTTI;

                    #endregion

                    if (sFormaRecepcion == "B")
                    {
                        #region Descarga XML

                        DownloadXML = serviceClienTFHKAReception.DescargarXML(ParametrosDownload);

                        if (DownloadXML.codigo == 200)
                        {
                            File.WriteAllBytes(sRutaArchivoXML, Convert.FromBase64String(DownloadXML.archivo));

                            string sSyncAttachmentCopy = sSyncAttachmentOrigin.Replace("%prmTipoAdjunto%", "XML").Replace("%pmrFile%", sRutaArchivoXML).Replace("%prmNIT%", U_BOTNI).Replace("%prmNF%", U_BOTNF).Replace("%CRWS%", DownloadXML.codigo.ToString());
                            oRsSyncDocument.DoQuery(sSyncAttachmentCopy);
                        }
                        else
                        {
                            string sSyncAttachmentCopy = sSyncAttachmentOrigin.Replace("%prmTipoAdjunto%", "XML").Replace("%pmrFile%", "").Replace("%prmNIT%", U_BOTNI).Replace("%prmNF%", "U_BOTNF").Replace("%CRWS%", DownloadXML.codigo.ToString());
                            oRsSyncDocument.DoQuery(sSyncAttachmentCopy);
                        }

                        #endregion
                    }
                    else if (sFormaRecepcion == "C")
                    {
                        #region Descarga XML

                        DownloadXML = serviceClienTFHKAReception.DescargarXML(ParametrosDownload);

                        if (DownloadXML.codigo == 200)
                        {
                            File.WriteAllBytes(sRutaArchivoXML, Convert.FromBase64String(DownloadXML.archivo));

                            string sSyncAttachmentCopy = sSyncAttachmentOrigin.Replace("%prmTipoAdjunto%", "XML").Replace("%pmrFile%", sRutaArchivoXML).Replace("%prmNIT%", U_BOTNI).Replace("%prmNF%", U_BOTNF).Replace("%CRWS%", DownloadXML.codigo.ToString());
                            oRsSyncDocument.DoQuery(sSyncAttachmentCopy);

                        }
                        else
                        {
                            string sSyncAttachmentCopy = sSyncAttachmentOrigin.Replace("%prmTipoAdjunto%", "XML").Replace("%pmrFile%", "").Replace("%prmNIT%", U_BOTNI).Replace("%prmNF%", "U_BOTNF").Replace("%CRWS%", DownloadXML.codigo.ToString());
                            oRsSyncDocument.DoQuery(sSyncAttachmentCopy);
                        }

                        #endregion

                        #region Descarga PDF

                        DownloadPDF = serviceClienTFHKAReception.DescargarRepGrafica(ParametrosDownload);

                        if (DownloadPDF.codigo == 200)
                        {
                            File.WriteAllBytes(sRutaArchivoPDF, Convert.FromBase64String(DownloadPDF.archivo));

                            string sSyncAttachmentCopy = sSyncAttachmentOrigin.Replace("%prmTipoAdjunto%", "PDF").Replace("%pmrFile%", sRutaArchivoPDF).Replace("%prmNIT%", U_BOTNI).Replace("%prmNF%", U_BOTNF).Replace("%CRWS%", DownloadPDF.codigo.ToString());
                            oRsSyncDocument.DoQuery(sSyncAttachmentCopy);
                        }
                        else
                        {
                            string sSyncAttachmentCopy = sSyncAttachmentOrigin.Replace("%prmTipoAdjunto%", "PDF").Replace("%pmrFile%", "").Replace("%prmNIT%", U_BOTNI).Replace("%prmNF%", U_BOTNF).Replace("%CRWS%", DownloadPDF.codigo.ToString());
                            oRsSyncDocument.DoQuery(sSyncAttachmentCopy);
                        }

                        #endregion

                    }

                    #endregion

                    #region Liberacion de Objetos                

                    DllFunciones.liberarObjetos(oRsSyncDocument);

                    #endregion
                }



            }
            catch (Exception)
            {
                throw;
            }

        }

        private void ChagueStatusDocument_TFHKA(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Form _oFormVisorRepcecion, ItemEvent pVal, string _ColUID)
        {

            if (_ColUID == "Col_15")
            {
                #region Acuse de Recibo

                SAPbouiCOM.Matrix oMatrixOPCH = (Matrix)_oFormVisorRepcecion.Items.Item("MtxOPCH").Specific;

                if (pVal.Row == 0)
                {

                }
                else
                {
                    #region Variables y Objetos

                    string sCondicionPago = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_24").Cells.Item(pVal.Row).Specific)).Value;
                    string sNumeroDocumentoFacturaProveedor = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific)).Value;
                    string sIdentificacionEmisor = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific)).Value;
                    string sTipoIdentificacionEmisor = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_14").Cells.Item(pVal.Row).Specific)).Value;
                    string sNombreProveedor = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_3").Cells.Item(pVal.Row).Specific)).Value;
                    string sEstadoDIAN = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific)).Value;

                    int iProcesar = 0;

                    string sCodigoEstadoDIAN = null;

                    #endregion

                    if (sCondicionPago == "Contado" || string.IsNullOrEmpty(sCondicionPago))
                    {

                    }
                    else if (sEstadoDIAN == "Recibo del bien y/o prestación del servicio")
                    {

                    }
                    else
                    {
                        #region Validacion estado documentos                        

                        if (sEstadoDIAN == "Cargado")
                        {
                            iProcesar = DllFunciones.sendMessageBoxY_N(_sboapp, "Se generara el Acuse de recibo (DIAN) en el documento " + sNumeroDocumentoFacturaProveedor + " del proveedor " + sNombreProveedor + " . ¿ Desea Continuar ? ");
                            sCodigoEstadoDIAN = "10";
                        }
                        else if (sEstadoDIAN == "Entregado")
                        {
                            iProcesar = DllFunciones.sendMessageBoxY_N(_sboapp, "Se generara el Acuse de recibo (DIAN) en el documento " + sNumeroDocumentoFacturaProveedor + " del proveedor " + sNombreProveedor + " . ¿ Desea Continuar ? ");
                            sCodigoEstadoDIAN = "10";
                        }
                        else if (sEstadoDIAN == "Precargado")
                        {
                            iProcesar = DllFunciones.sendMessageBoxY_N(_sboapp, "Se generara el Acuse de recibo (DIAN) en el documento " + sNumeroDocumentoFacturaProveedor + " del proveedor " + sNombreProveedor + " . ¿ Desea Continuar ? ");
                            sCodigoEstadoDIAN = "10";
                        }
                        else if (sEstadoDIAN == "Acuse de recibo (DIAN)")
                        {
                            iProcesar = 0;
                        }
                        else
                        {
                            DllFunciones.sendMessageBox(_sboapp, "No se puede generar el Acuse de recibo (DIAN) en el documento " + sNumeroDocumentoFacturaProveedor + " del proveedor " + sNombreProveedor + " . Por favor validar con su administrador.");

                            iProcesar = 0;
                        }

                        #endregion

                        if (iProcesar == 1)
                        {
                            #region Procesa Evento DIAN

                            #region Variables y Objetos 

                            string sNombreAceptador = string.Empty;
                            string sApellidoAceptador = string.Empty;
                            string sCargoAceptador = string.Empty;
                            string sDepartamentoAceptador = string.Empty;
                            string sNITAceptador = string.Empty;
                            string sTipoDocumentoAceptador = string.Empty;
                            string sDigitoVerificacionAceptador = string.Empty;
                            string UsuarioSAPActual = string.Empty;
                            string sGetauthorizer = string.Empty;

                            UsuarioSAPActual = Convert.ToString(_oCompany.UserSignature);

                            #endregion

                            #region Valida si esta configurado el usuario                   

                            SAPbobsCOM.Recordset oGetauthorizer = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                            sGetauthorizer = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "Getauthorizer");

                            sGetauthorizer = sGetauthorizer.Replace("%UserId%", UsuarioSAPActual);

                            oGetauthorizer.DoQuery(sGetauthorizer);

                            #endregion

                            if (oGetauthorizer.RecordCount > 0)
                            {
                                sNombreAceptador = Convert.ToString(oGetauthorizer.Fields.Item("Nombre").Value.ToString());
                                sApellidoAceptador = Convert.ToString(oGetauthorizer.Fields.Item("Apellido").Value.ToString());
                                sCargoAceptador = Convert.ToString(oGetauthorizer.Fields.Item("Cargo").Value.ToString());
                                sDepartamentoAceptador = Convert.ToString(oGetauthorizer.Fields.Item("Departamento").Value.ToString());
                                sNITAceptador = Convert.ToString(oGetauthorizer.Fields.Item("NIT").Value.ToString());
                                sTipoDocumentoAceptador = Convert.ToString(oGetauthorizer.Fields.Item("TipoDocumento").Value.ToString());

                                #region Valida campos obligatorios 

                                if (string.IsNullOrEmpty(sNombreAceptador))
                                {
                                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10028", _sboapp, null));
                                }
                                else if (string.IsNullOrEmpty(sApellidoAceptador))
                                {
                                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10029", _sboapp, null));
                                }
                                else if (string.IsNullOrEmpty(sCargoAceptador))
                                {
                                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10030", _sboapp, null));
                                }
                                else if (string.IsNullOrEmpty(sDepartamentoAceptador))
                                {
                                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10031", _sboapp, null));
                                }
                                else if (string.IsNullOrEmpty(sNITAceptador))
                                {
                                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10032", _sboapp, null));
                                }
                                else if (string.IsNullOrEmpty(sTipoDocumentoAceptador))
                                {
                                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10033", _sboapp, null));
                                }
                                else
                                {

                                    #region Consulta URL

                                    string sGetModo = null;
                                    string sURLRecepcion = null;
                                    string sModo = null;
                                    string sRutaXML = null;
                                    string sRutaPDF = null;
                                    string sProtocoloComunicacion = null;
                                    string sTokenEmpresa = null;
                                    string sTokenPassword = null;
                                    string sGetDV = string.Empty;


                                    SAPbobsCOM.Recordset oConsultarGetModo = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                    sGetModo = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetModoandURL");

                                    sGetModo = sGetModo.Replace("%Estado%", "\"U_BO_Status\" = 'Y'").Replace("%DocEntry%", " ");

                                    oConsultarGetModo.DoQuery(sGetModo);

                                    sURLRecepcion = Convert.ToString(oConsultarGetModo.Fields.Item("URLTFHKA").Value.ToString()) + "/ReceptorWS.svc?wsdl";
                                    sModo = Convert.ToString(oConsultarGetModo.Fields.Item("Modo").Value.ToString());

                                    sProtocoloComunicacion = Convert.ToString(oConsultarGetModo.Fields.Item("ProtocoloComunicacion").Value.ToString());
                                    sTokenEmpresa = Convert.ToString(oConsultarGetModo.Fields.Item("TokenEmpresa").Value.ToString());
                                    sTokenPassword = Convert.ToString(oConsultarGetModo.Fields.Item("TokenPassword").Value.ToString());
                                    sRutaXML = Convert.ToString(oConsultarGetModo.Fields.Item("RutaXML").Value.ToString());
                                    sRutaPDF = Convert.ToString(oConsultarGetModo.Fields.Item("RutaPDF").Value.ToString());

                                    DllFunciones.liberarObjetos(oConsultarGetModo);

                                    #endregion

                                    #region Instanciacion parametros TFHKA

                                    if (sProtocoloComunicacion == "HTTP")
                                    {
                                        BasicHttpBinding port = new BasicHttpBinding();
                                    }
                                    else if (sProtocoloComunicacion == "HTTPS")
                                    {
                                        BasicHttpsBinding port = new BasicHttpsBinding();
                                    }

                                    port.MaxBufferPoolSize = Int32.MaxValue;
                                    port.MaxBufferSize = Int32.MaxValue;
                                    port.MaxReceivedMessageSize = Int32.MaxValue;
                                    port.ReaderQuotas.MaxStringContentLength = Int32.MaxValue;
                                    port.SendTimeout = TimeSpan.FromMinutes(2);
                                    port.ReceiveTimeout = TimeSpan.FromMinutes(2);

                                    if (sProtocoloComunicacion == "HTTPS")
                                    {
                                        port.Security.Mode = BasicHttpSecurityMode.Transport;
                                    }

                                    EndpointAddress endPointEmision = new EndpointAddress(sURLRecepcion); //URL

                                    Recepcion21WS.ReceptorWSClient serviceClienTFHKAReception;
                                    serviceClienTFHKAReception = new Recepcion21WS.ReceptorWSClient(port, endPointEmision);

                                    #endregion

                                    #region Parametros generales Reporte

                                    ReceptorCambioEstatusRequest ParametrosCambioEstatus = new ReceptorCambioEstatusRequest();

                                    ParametrosCambioEstatus.tokenEmpresa = sTokenEmpresa;
                                    ParametrosCambioEstatus.tokenPassword = sTokenPassword;

                                    ParametrosCambioEstatus.identificadorEmisor = sIdentificacionEmisor;
                                    ParametrosCambioEstatus.tipoIdentificacionemisor = sTipoIdentificacionEmisor;
                                    ParametrosCambioEstatus.numeroDocumento = sNumeroDocumentoFacturaProveedor;

                                    ReceptorCambioEstatusRequest.EjecutadoPorRequest UsuarioAceptador = new ReceptorCambioEstatusRequest.EjecutadoPorRequest();

                                    UsuarioAceptador.Nombre = sNombreAceptador;
                                    UsuarioAceptador.Apellido = sApellidoAceptador;
                                    UsuarioAceptador.Cargo = sCargoAceptador;
                                    UsuarioAceptador.Departamento = sDepartamentoAceptador;

                                    ParametrosCambioEstatus.EjecutadoPor = UsuarioAceptador;

                                    ReceptorCambioEstatusRequest.EjecutadoPorRequest.IdentificacionRequest NITAceptador = new ReceptorCambioEstatusRequest.EjecutadoPorRequest.IdentificacionRequest();

                                    NITAceptador.NumeroIdentificacion = sNITAceptador;
                                    NITAceptador.TipoIdentificacion = sTipoDocumentoAceptador;

                                    #region Consulta Digito Verificacion                   

                                    SAPbobsCOM.Recordset oGetDV = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                    sGetDV = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetDV");

                                    sGetDV = sGetDV.Replace("%NIT%", sNITAceptador);

                                    oGetDV.DoQuery(sGetDV);

                                    if (oGetDV.RecordCount > 0)
                                    {
                                        sDigitoVerificacionAceptador = Convert.ToString(oGetDV.Fields.Item("DV").Value.ToString());
                                    }

                                    #endregion

                                    NITAceptador.Dv = sDigitoVerificacionAceptador;

                                    ParametrosCambioEstatus.EjecutadoPor.Identificacion = NITAceptador;

                                    ParametrosCambioEstatus.status = sCodigoEstadoDIAN;
                                    ParametrosCambioEstatus.codigoRechazo = "02";

                                    #endregion

                                    #region Cambia estado en DIAN                            

                                    Recepcion21WS.ResponseGeneral wsCambiarEstado;

                                    wsCambiarEstado = serviceClienTFHKAReception.CambioEstatus(ParametrosCambioEstatus);

                                    if (wsCambiarEstado.codigo == 200)
                                    {
                                        #region Actualiza estado documento en SAP

                                        SAPbobsCOM.Recordset oUpdateStatusDocument = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                        string sUpdateStatusDocument = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "PostUpdateStatusDocument");

                                        sUpdateStatusDocument = sUpdateStatusDocument.Replace("%NumeroFactura%", sNumeroDocumentoFacturaProveedor).Replace("%NumeroIdentificacion%", sIdentificacionEmisor).Replace("%CodigoEventoDIANPT%", sCodigoEstadoDIAN);

                                        oUpdateStatusDocument.DoQuery(sUpdateStatusDocument);

                                        DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10034", _sboapp, null));

                                        #endregion
                                    }
                                    else
                                    {
                                        DllFunciones.sendMessageBox(_sboapp, "RE007 - No se pudo cambiar el estaodo" + wsCambiarEstado.mensaje.ToString());
                                    }

                                    #endregion

                                }

                                #endregion


                            }
                            else
                            {
                                DllFunciones.sendMessageBox(_sboapp, "El usuario en SAP no esta autorizado para generar eventos en las facturas de proveedor ");
                            }

                            #endregion
                        }
                    }
                }

                #endregion
            }
            if (_ColUID == "Col_27")
            {
                #region Recibo del bien o servicio

                SAPbouiCOM.Matrix oMatrixOPCH = (Matrix)_oFormVisorRepcecion.Items.Item("MtxOPCH").Specific;

                if (pVal.Row == 0)
                {

                }
                else
                {
                    #region Variables y Objetos

                    string sCondicionPago = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_24").Cells.Item(pVal.Row).Specific)).Value;
                    string sNumeroDocumentoFacturaProveedor = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific)).Value;
                    string sIdentificacionEmisor = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific)).Value;
                    string sTipoIdentificacionEmisor = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_14").Cells.Item(pVal.Row).Specific)).Value;
                    string sNombreProveedor = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_3").Cells.Item(pVal.Row).Specific)).Value;
                    string sEstadoDIAN = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific)).Value;

                    int iProcesar = 0;

                    string sCodigoEstadoDIAN = null;

                    #endregion

                    if (sCondicionPago == "Contado" || string.IsNullOrEmpty(sCondicionPago))
                    {

                    }
                    else
                    {
                        #region Validacion estado documentos                        

                        if (sEstadoDIAN == "Acuse de recibo (DIAN)")
                        {
                            iProcesar = DllFunciones.sendMessageBoxY_N(_sboapp, "Se generara el Recibo del bien y/o prestación del servicio en el documento " + sNumeroDocumentoFacturaProveedor + " del proveedor " + sNombreProveedor + " . ¿ Desea Continuar ? ");
                            sCodigoEstadoDIAN = "12";
                        }
                        else 
                        {
                            iProcesar = 0;
                        }

                        #endregion

                        if (iProcesar == 1)
                        {
                            #region Procesa Evento DIAN

                            #region Variables y Objetos 

                            string sNombreAceptador = string.Empty;
                            string sApellidoAceptador = string.Empty;
                            string sCargoAceptador = string.Empty;
                            string sDepartamentoAceptador = string.Empty;
                            string sNITAceptador = string.Empty;
                            string sTipoDocumentoAceptador = string.Empty;
                            string sDigitoVerificacionAceptador = string.Empty;
                            string UsuarioSAPActual = string.Empty;
                            string sGetauthorizer = string.Empty;

                            UsuarioSAPActual = Convert.ToString(_oCompany.UserSignature);

                            #endregion

                            #region Valida si esta configurado el usuario                   

                            SAPbobsCOM.Recordset oGetauthorizer = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                            sGetauthorizer = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "Getauthorizer");

                            sGetauthorizer = sGetauthorizer.Replace("%UserId%", UsuarioSAPActual);

                            oGetauthorizer.DoQuery(sGetauthorizer);

                            #endregion

                            if (oGetauthorizer.RecordCount > 0)
                            {
                                sNombreAceptador = Convert.ToString(oGetauthorizer.Fields.Item("Nombre").Value.ToString());
                                sApellidoAceptador = Convert.ToString(oGetauthorizer.Fields.Item("Apellido").Value.ToString());
                                sCargoAceptador = Convert.ToString(oGetauthorizer.Fields.Item("Cargo").Value.ToString());
                                sDepartamentoAceptador = Convert.ToString(oGetauthorizer.Fields.Item("Departamento").Value.ToString());
                                sNITAceptador = Convert.ToString(oGetauthorizer.Fields.Item("NIT").Value.ToString());
                                sTipoDocumentoAceptador = Convert.ToString(oGetauthorizer.Fields.Item("TipoDocumento").Value.ToString());

                                #region Valida campos obligatorios 

                                if (string.IsNullOrEmpty(sNombreAceptador))
                                {
                                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10028", _sboapp, null));
                                }
                                else if (string.IsNullOrEmpty(sApellidoAceptador))
                                {
                                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10029", _sboapp, null));
                                }
                                else if (string.IsNullOrEmpty(sCargoAceptador))
                                {
                                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10030", _sboapp, null));
                                }
                                else if (string.IsNullOrEmpty(sDepartamentoAceptador))
                                {
                                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10031", _sboapp, null));
                                }
                                else if (string.IsNullOrEmpty(sNITAceptador))
                                {
                                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10032", _sboapp, null));
                                }
                                else if (string.IsNullOrEmpty(sTipoDocumentoAceptador))
                                {
                                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10033", _sboapp, null));
                                }
                                else
                                {

                                    #region Consulta URL

                                    string sGetModo = null;
                                    string sURLRecepcion = null;
                                    string sModo = null;
                                    string sRutaXML = null;
                                    string sRutaPDF = null;
                                    string sProtocoloComunicacion = null;
                                    string sTokenEmpresa = null;
                                    string sTokenPassword = null;
                                    string sGetDV = string.Empty;


                                    SAPbobsCOM.Recordset oConsultarGetModo = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                    sGetModo = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetModoandURL");

                                    sGetModo = sGetModo.Replace("%Estado%", "\"U_BO_Status\" = 'Y'").Replace("%DocEntry%", " ");

                                    oConsultarGetModo.DoQuery(sGetModo);

                                    sURLRecepcion = Convert.ToString(oConsultarGetModo.Fields.Item("URLTFHKA").Value.ToString()) + "/ReceptorWS.svc?wsdl";
                                    sModo = Convert.ToString(oConsultarGetModo.Fields.Item("Modo").Value.ToString());

                                    sProtocoloComunicacion = Convert.ToString(oConsultarGetModo.Fields.Item("ProtocoloComunicacion").Value.ToString());
                                    sTokenEmpresa = Convert.ToString(oConsultarGetModo.Fields.Item("TokenEmpresa").Value.ToString());
                                    sTokenPassword = Convert.ToString(oConsultarGetModo.Fields.Item("TokenPassword").Value.ToString());
                                    sRutaXML = Convert.ToString(oConsultarGetModo.Fields.Item("RutaXML").Value.ToString());
                                    sRutaPDF = Convert.ToString(oConsultarGetModo.Fields.Item("RutaPDF").Value.ToString());

                                    DllFunciones.liberarObjetos(oConsultarGetModo);

                                    #endregion

                                    #region Instanciacion parametros TFHKA

                                    if (sProtocoloComunicacion == "HTTP")
                                    {
                                        BasicHttpBinding port = new BasicHttpBinding();
                                    }
                                    else if (sProtocoloComunicacion == "HTTPS")
                                    {
                                        BasicHttpsBinding port = new BasicHttpsBinding();
                                    }

                                    port.MaxBufferPoolSize = Int32.MaxValue;
                                    port.MaxBufferSize = Int32.MaxValue;
                                    port.MaxReceivedMessageSize = Int32.MaxValue;
                                    port.ReaderQuotas.MaxStringContentLength = Int32.MaxValue;
                                    port.SendTimeout = TimeSpan.FromMinutes(2);
                                    port.ReceiveTimeout = TimeSpan.FromMinutes(2);

                                    if (sProtocoloComunicacion == "HTTPS")
                                    {
                                        port.Security.Mode = BasicHttpSecurityMode.Transport;
                                    }

                                    EndpointAddress endPointEmision = new EndpointAddress(sURLRecepcion); //URL

                                    Recepcion21WS.ReceptorWSClient serviceClienTFHKAReception;
                                    serviceClienTFHKAReception = new Recepcion21WS.ReceptorWSClient(port, endPointEmision);

                                    #endregion

                                    #region Parametros generales Reporte

                                    ReceptorCambioEstatusRequest ParametrosCambioEstatus = new ReceptorCambioEstatusRequest();

                                    ParametrosCambioEstatus.tokenEmpresa = sTokenEmpresa;
                                    ParametrosCambioEstatus.tokenPassword = sTokenPassword;

                                    ParametrosCambioEstatus.identificadorEmisor = sIdentificacionEmisor;
                                    ParametrosCambioEstatus.tipoIdentificacionemisor = sTipoIdentificacionEmisor;
                                    ParametrosCambioEstatus.numeroDocumento = sNumeroDocumentoFacturaProveedor;

                                    ReceptorCambioEstatusRequest.EjecutadoPorRequest UsuarioAceptador = new ReceptorCambioEstatusRequest.EjecutadoPorRequest();

                                    UsuarioAceptador.Nombre = sNombreAceptador;
                                    UsuarioAceptador.Apellido = sApellidoAceptador;
                                    UsuarioAceptador.Cargo = sCargoAceptador;
                                    UsuarioAceptador.Departamento = sDepartamentoAceptador;

                                    ParametrosCambioEstatus.EjecutadoPor = UsuarioAceptador;

                                    ReceptorCambioEstatusRequest.EjecutadoPorRequest.IdentificacionRequest NITAceptador = new ReceptorCambioEstatusRequest.EjecutadoPorRequest.IdentificacionRequest();

                                    NITAceptador.NumeroIdentificacion = sNITAceptador;
                                    NITAceptador.TipoIdentificacion = sTipoDocumentoAceptador;

                                    #region Consulta Digito Verificacion                   

                                    SAPbobsCOM.Recordset oGetDV = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                    sGetDV = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetDV");

                                    sGetDV = sGetDV.Replace("%NIT%", sNITAceptador);

                                    oGetDV.DoQuery(sGetDV);

                                    if (oGetDV.RecordCount > 0)
                                    {
                                        sDigitoVerificacionAceptador = Convert.ToString(oGetDV.Fields.Item("DV").Value.ToString());
                                    }

                                    #endregion

                                    NITAceptador.Dv = sDigitoVerificacionAceptador;

                                    ParametrosCambioEstatus.EjecutadoPor.Identificacion = NITAceptador;

                                    ParametrosCambioEstatus.status = sCodigoEstadoDIAN;
                                    ParametrosCambioEstatus.codigoRechazo = "02";

                                    #endregion

                                    #region Cambia estado en DIAN                            

                                    Recepcion21WS.ResponseGeneral wsCambiarEstado;

                                    wsCambiarEstado = serviceClienTFHKAReception.CambioEstatus(ParametrosCambioEstatus);

                                    if (wsCambiarEstado.codigo == 200)
                                    {
                                        #region Actualiza estado documento en SAP

                                        SAPbobsCOM.Recordset oUpdateStatusDocument = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                        string sUpdateStatusDocument = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "PostUpdateStatusDocument");

                                        sUpdateStatusDocument = sUpdateStatusDocument.Replace("%NumeroFactura%", sNumeroDocumentoFacturaProveedor).Replace("%NumeroIdentificacion%", sIdentificacionEmisor).Replace("%CodigoEventoDIANPT%", sCodigoEstadoDIAN);

                                        oUpdateStatusDocument.DoQuery(sUpdateStatusDocument);

                                        DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10034", _sboapp, null));

                                        #endregion
                                    }
                                    else
                                    {
                                        DllFunciones.sendMessageBox(_sboapp, "RE007 - No se pudo cambiar el estaodo" + wsCambiarEstado.mensaje.ToString());
                                    }

                                    #endregion

                                }

                                #endregion


                            }
                            else
                            {
                                DllFunciones.sendMessageBox(_sboapp, "El usuario en SAP no esta autorizado para generar eventos en las facturas de proveedor ");
                            }

                            #endregion
                        }
                    }
                }

                #endregion
            }
            else if (_ColUID == "Col_22")
            {
                #region Aceptacion documento

                SAPbouiCOM.Matrix oMatrixOPCH = (Matrix)_oFormVisorRepcecion.Items.Item("MtxOPCH").Specific;

                if (pVal.Row == 0)
                {

                }
                else
                {
                    #region Variables y Objetos

                    string sCondicionPago = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_24").Cells.Item(pVal.Row).Specific)).Value;
                    string sNumeroDocumentoFacturaProveedor = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific)).Value;
                    string sIdentificacionEmisor = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific)).Value;
                    string sTipoIdentificacionEmisor = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_14").Cells.Item(pVal.Row).Specific)).Value;
                    string sNombreProveedor = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_3").Cells.Item(pVal.Row).Specific)).Value;
                    string sEstadoDIAN = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific)).Value;

                    int iProcesar = 0;
                    string sCodigoEstadoDIAN = null;

                    #endregion

                    if (sCondicionPago == "Contado" || string.IsNullOrEmpty(sCondicionPago))
                    {

                    }
                    else
                    {
                        #region Validacion estado documentos                        

                        if (sEstadoDIAN == "Cargado")
                        {
                            iProcesar = DllFunciones.sendMessageBoxY_N(_sboapp, "Se generara el Acuse de recibo (DIAN) en el documento " + sNumeroDocumentoFacturaProveedor + " del proveedor " + sNombreProveedor + " . ¿ Desea Continuar ? ");
                            sCodigoEstadoDIAN = "10";
                        }
                        else if (sEstadoDIAN == "Entregado")
                        {
                            iProcesar = DllFunciones.sendMessageBoxY_N(_sboapp, "Se generara el Acuse de recibo (DIAN) en el documento " + sNumeroDocumentoFacturaProveedor + " del proveedor " + sNombreProveedor + " . ¿ Desea Continuar ? ");
                            sCodigoEstadoDIAN = "10";
                        }
                        else if (sEstadoDIAN == "Precargado")
                        {
                            iProcesar = DllFunciones.sendMessageBoxY_N(_sboapp, "Se generara el Acuse de recibo (DIAN) en el documento " + sNumeroDocumentoFacturaProveedor + " del proveedor " + sNombreProveedor + " . ¿ Desea Continuar ? ");
                            sCodigoEstadoDIAN = "10";
                        }
                        else if (sEstadoDIAN == "Acuse de recibo (DIAN)")
                        {
                            iProcesar = DllFunciones.sendMessageBoxY_N(_sboapp, "Se generara el Recibo del bien y/o prestación del servicio en el documento " + sNumeroDocumentoFacturaProveedor + " del proveedor " + sNombreProveedor + " . ¿ Desea Continuar ? ");
                            sCodigoEstadoDIAN = "12";
                        }
                        else if (sEstadoDIAN == "Recibo del bien y/o prestación del servicio")
                        {
                            iProcesar = DllFunciones.sendMessageBoxY_N(_sboapp, "Se generara la Aceptación expresa (DIAN) en el documento " + sNumeroDocumentoFacturaProveedor + " del proveedor " + sNombreProveedor + " . ¿ Desea Continuar ? ");
                            sCodigoEstadoDIAN = "02";
                        }
                        else if (sEstadoDIAN == "Aceptación expresa (DIAN)")
                        {
                            DllFunciones.sendMessageBox(_sboapp, "El documento ya se encuentra con Aceptación expresa (DIAN) en el documento " + sNumeroDocumentoFacturaProveedor + " del proveedor " + sNombreProveedor + " . No se puede cambiar el estado");

                            iProcesar = 0;
                        }

                        #endregion

                        if (iProcesar == 1)
                        {
                            #region Procesa Evento DIAN

                            #region Variables y Objetos 

                            string sNombreAceptador = string.Empty;
                            string sApellidoAceptador = string.Empty;
                            string sCargoAceptador = string.Empty;
                            string sDepartamentoAceptador = string.Empty;
                            string sNITAceptador = string.Empty;
                            string sTipoDocumentoAceptador = string.Empty;
                            string sDigitoVerificacionAceptador = string.Empty;
                            string UsuarioSAPActual = string.Empty;
                            string sGetauthorizer = string.Empty;

                            UsuarioSAPActual = Convert.ToString(_oCompany.UserSignature);

                            #endregion

                            #region Valida si esta configurado el usuario                   

                            SAPbobsCOM.Recordset oGetauthorizer = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                            sGetauthorizer = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "Getauthorizer");

                            sGetauthorizer = sGetauthorizer.Replace("%UserId%", UsuarioSAPActual);

                            oGetauthorizer.DoQuery(sGetauthorizer);

                            #endregion

                            if (oGetauthorizer.RecordCount > 0)
                            {
                                sNombreAceptador = Convert.ToString(oGetauthorizer.Fields.Item("Nombre").Value.ToString());
                                sApellidoAceptador = Convert.ToString(oGetauthorizer.Fields.Item("Apellido").Value.ToString());
                                sCargoAceptador = Convert.ToString(oGetauthorizer.Fields.Item("Cargo").Value.ToString());
                                sDepartamentoAceptador = Convert.ToString(oGetauthorizer.Fields.Item("Departamento").Value.ToString());
                                sNITAceptador = Convert.ToString(oGetauthorizer.Fields.Item("NIT").Value.ToString());
                                sTipoDocumentoAceptador = Convert.ToString(oGetauthorizer.Fields.Item("TipoDocumento").Value.ToString());

                                #region Valida campos obligatorios 

                                if (string.IsNullOrEmpty(sNombreAceptador))
                                {
                                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10028", _sboapp, null));
                                }
                                else if (string.IsNullOrEmpty(sApellidoAceptador))
                                {
                                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10029", _sboapp, null));
                                }
                                else if (string.IsNullOrEmpty(sCargoAceptador))
                                {
                                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10030", _sboapp, null));
                                }
                                else if (string.IsNullOrEmpty(sDepartamentoAceptador))
                                {
                                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10031", _sboapp, null));
                                }
                                else if (string.IsNullOrEmpty(sNITAceptador))
                                {
                                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10032", _sboapp, null));
                                }
                                else if (string.IsNullOrEmpty(sTipoDocumentoAceptador))
                                {
                                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10033", _sboapp, null));
                                }
                                else
                                {

                                    #region Consulta URL

                                    string sGetModo = null;
                                    string sURLRecepcion = null;
                                    string sModo = null;
                                    string sRutaXML = null;
                                    string sRutaPDF = null;
                                    string sProtocoloComunicacion = null;
                                    string sTokenEmpresa = null;
                                    string sTokenPassword = null;
                                    string sGetDV = string.Empty;


                                    SAPbobsCOM.Recordset oConsultarGetModo = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                    sGetModo = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetModoandURL");

                                    sGetModo = sGetModo.Replace("%Estado%", "\"U_BO_Status\" = 'Y'").Replace("%DocEntry%", " ");

                                    oConsultarGetModo.DoQuery(sGetModo);

                                    sURLRecepcion = Convert.ToString(oConsultarGetModo.Fields.Item("URLTFHKA").Value.ToString()) + "/ReceptorWS.svc?wsdl";
                                    sModo = Convert.ToString(oConsultarGetModo.Fields.Item("Modo").Value.ToString());

                                    sProtocoloComunicacion = Convert.ToString(oConsultarGetModo.Fields.Item("ProtocoloComunicacion").Value.ToString());
                                    sTokenEmpresa = Convert.ToString(oConsultarGetModo.Fields.Item("TokenEmpresa").Value.ToString());
                                    sTokenPassword = Convert.ToString(oConsultarGetModo.Fields.Item("TokenPassword").Value.ToString());
                                    sRutaXML = Convert.ToString(oConsultarGetModo.Fields.Item("RutaXML").Value.ToString());
                                    sRutaPDF = Convert.ToString(oConsultarGetModo.Fields.Item("RutaPDF").Value.ToString());

                                    DllFunciones.liberarObjetos(oConsultarGetModo);

                                    #endregion

                                    #region Instanciacion parametros TFHKA

                                    if (sProtocoloComunicacion == "HTTP")
                                    {
                                        BasicHttpBinding port = new BasicHttpBinding();
                                    }
                                    else if (sProtocoloComunicacion == "HTTPS")
                                    {
                                        BasicHttpsBinding port = new BasicHttpsBinding();
                                    }

                                    port.MaxBufferPoolSize = Int32.MaxValue;
                                    port.MaxBufferSize = Int32.MaxValue;
                                    port.MaxReceivedMessageSize = Int32.MaxValue;
                                    port.ReaderQuotas.MaxStringContentLength = Int32.MaxValue;
                                    port.SendTimeout = TimeSpan.FromMinutes(2);
                                    port.ReceiveTimeout = TimeSpan.FromMinutes(2);

                                    if (sProtocoloComunicacion == "HTTPS")
                                    {
                                        port.Security.Mode = BasicHttpSecurityMode.Transport;
                                    }

                                    EndpointAddress endPointEmision = new EndpointAddress(sURLRecepcion); //URL

                                    Recepcion21WS.ReceptorWSClient serviceClienTFHKAReception;
                                    serviceClienTFHKAReception = new Recepcion21WS.ReceptorWSClient(port, endPointEmision);

                                    #endregion

                                    #region Parametros generales Reporte

                                    ReceptorCambioEstatusRequest ParametrosCambioEstatus = new ReceptorCambioEstatusRequest();

                                    ParametrosCambioEstatus.tokenEmpresa = sTokenEmpresa;
                                    ParametrosCambioEstatus.tokenPassword = sTokenPassword;

                                    ParametrosCambioEstatus.identificadorEmisor = sIdentificacionEmisor;
                                    ParametrosCambioEstatus.tipoIdentificacionemisor = sTipoIdentificacionEmisor;
                                    ParametrosCambioEstatus.numeroDocumento = sNumeroDocumentoFacturaProveedor;

                                    ReceptorCambioEstatusRequest.EjecutadoPorRequest UsuarioAceptador = new ReceptorCambioEstatusRequest.EjecutadoPorRequest();

                                    UsuarioAceptador.Nombre = sNombreAceptador;
                                    UsuarioAceptador.Apellido = sApellidoAceptador;
                                    UsuarioAceptador.Cargo = sCargoAceptador;
                                    UsuarioAceptador.Departamento = sDepartamentoAceptador;

                                    ParametrosCambioEstatus.EjecutadoPor = UsuarioAceptador;

                                    ReceptorCambioEstatusRequest.EjecutadoPorRequest.IdentificacionRequest NITAceptador = new ReceptorCambioEstatusRequest.EjecutadoPorRequest.IdentificacionRequest();

                                    NITAceptador.NumeroIdentificacion = sNITAceptador;
                                    NITAceptador.TipoIdentificacion = sTipoDocumentoAceptador;

                                    #region Consulta Digito Verificacion                   

                                    SAPbobsCOM.Recordset oGetDV = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                    sGetDV = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetDV");

                                    sGetDV = sGetDV.Replace("%NIT%", sNITAceptador);

                                    oGetDV.DoQuery(sGetDV);

                                    if (oGetDV.RecordCount > 0)
                                    {
                                        sDigitoVerificacionAceptador = Convert.ToString(oGetDV.Fields.Item("DV").Value.ToString());
                                    }

                                    #endregion

                                    NITAceptador.Dv = sDigitoVerificacionAceptador;

                                    ParametrosCambioEstatus.EjecutadoPor.Identificacion = NITAceptador;

                                    ParametrosCambioEstatus.status = sCodigoEstadoDIAN;
                                    ParametrosCambioEstatus.codigoRechazo = "02";

                                    #endregion

                                    #region Cambia estado en DIAN                            

                                    Recepcion21WS.ResponseGeneral wsCambiarEstado;

                                    wsCambiarEstado = serviceClienTFHKAReception.CambioEstatus(ParametrosCambioEstatus);

                                    if (wsCambiarEstado.codigo == 200)
                                    {
                                        #region Actualiza estado documento en SAP

                                        SAPbobsCOM.Recordset oUpdateStatusDocument = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                        string sUpdateStatusDocument = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "PostUpdateStatusDocument");

                                        sUpdateStatusDocument = sUpdateStatusDocument.Replace("%NumeroFactura%", sNumeroDocumentoFacturaProveedor).Replace("%NumeroIdentificacion%", sIdentificacionEmisor).Replace("%CodigoEventoDIANPT%", sCodigoEstadoDIAN);

                                        oUpdateStatusDocument.DoQuery(sUpdateStatusDocument);

                                        DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10034", _sboapp, null));

                                        #endregion
                                    }
                                    else
                                    {
                                        DllFunciones.sendMessageBox(_sboapp, "RE007 - No se pudo cambiar el estaodo" + wsCambiarEstado.mensaje.ToString());
                                    }

                                    #endregion

                                }

                                #endregion


                            }
                            else
                            {
                                DllFunciones.sendMessageBox(_sboapp, "El usuario en SAP no esta autorizado para generar eventos en las facturas de proveedor ");
                            }

                            #endregion
                        }
                    }
                }

                #endregion
            }
            else if (_ColUID == "Col_7")
            {
                #region Rechazo del  documento

                SAPbouiCOM.Matrix oMatrixOPCH = (Matrix)_oFormVisorRepcecion.Items.Item("MtxOPCH").Specific;

                if (pVal.Row == 0)
                {

                }
                else
                {
                    #region Variables y Objetos

                    string sCondicionPago = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_24").Cells.Item(pVal.Row).Specific)).Value;                    
                    string sNumeroDocumentoFacturaProveedor = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific)).Value;
                    string sIdentificacionEmisor = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific)).Value;
                    string sTipoIdentificacionEmisor = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_14").Cells.Item(pVal.Row).Specific)).Value;
                    string sNombreProveedor = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_3").Cells.Item(pVal.Row).Specific)).Value;
                    string sEstadoDIAN = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific)).Value;

                    int iProcesar = 0;
                    string sCodigoEstadoDIAN = null;

                    #endregion

                    if (sCondicionPago == "Contado" || string.IsNullOrEmpty(sCondicionPago))
                    {

                    }
                    else if (sEstadoDIAN == "Aceptación expresa (DIAN)")
                    {

                    }
                    else
                    {
                        #region Validacion estados del documento

                        if (sEstadoDIAN == "Cargado")
                        {
                            iProcesar = DllFunciones.sendMessageBoxY_N(_sboapp, "Se generara el Acuse de recibo (DIAN) en el documento " + sNumeroDocumentoFacturaProveedor + " del proveedor " + sNombreProveedor + " para poder realizar el rechazo. ¿ Desea Continuar ? ");
                            sCodigoEstadoDIAN = "10";
                        }
                        else if (sEstadoDIAN == "Entregado")
                        {
                            iProcesar = DllFunciones.sendMessageBoxY_N(_sboapp, "Se generara el Acuse de recibo (DIAN) en el documento " + sNumeroDocumentoFacturaProveedor + " del proveedor " + sNombreProveedor + " para poder realizar el rechazo. ¿ Desea Continuar ? ");
                            sCodigoEstadoDIAN = "10";
                        }
                        else if (sEstadoDIAN == "Precargado")
                        {
                            iProcesar = DllFunciones.sendMessageBoxY_N(_sboapp, "Se generara el Acuse de recibo (DIAN) en el documento " + sNumeroDocumentoFacturaProveedor + " del proveedor " + sNombreProveedor + " para poder realizar el rechazo. ¿ Desea Continuar ? ");
                            sCodigoEstadoDIAN = "01";
                        }
                        else if (sEstadoDIAN == "Acuse de recibo (DIAN)")
                        {
                            iProcesar = DllFunciones.sendMessageBoxY_N(_sboapp, "Se generara el Recibo del bien y/o prestación del servicio en el documento " + sNumeroDocumentoFacturaProveedor + " del proveedor " + sNombreProveedor + " para poder realizar el rechazo. ¿ Desea continuar ?");
                            sCodigoEstadoDIAN = "12";
                        }
                        else if (sEstadoDIAN == "Recibo del bien y/o prestación del servicio")
                        {
                            iProcesar = DllFunciones.sendMessageBoxY_N(_sboapp, "Se rechazara el documento " + sNumeroDocumentoFacturaProveedor + " del proveedor " + sNombreProveedor + " . ¿ Desea continuar ?");
                            //DllFunciones.sendMessageBox(_sboapp, "No se puede rechazar el documento " + sNumeroDocumentoFacturaProveedor + " del proveedor " + sNombreProveedor + " . Ya se le genero - Recibo del bien y/o prestación del servicio. ");
                            sCodigoEstadoDIAN = "04";
                        }
                        else if (sEstadoDIAN == "Aceptación expresa (DIAN)")
                        {
                            DllFunciones.sendMessageBox(_sboapp, "El documento ya se encuentra con Aceptación expresa (DIAN) en el documento " + sNumeroDocumentoFacturaProveedor + " del proveedor " + sNombreProveedor + " . No se puede cambiar el estado");
                            iProcesar = 0;
                        }

                        #endregion

                        if (iProcesar == 1)
                        {
                            #region Variables y Objetos 

                            string sNombreAceptador = string.Empty;
                            string sApellidoAceptador = string.Empty;
                            string sCargoAceptador = string.Empty;
                            string sDepartamentoAceptador = string.Empty;
                            string sNITAceptador = string.Empty;
                            string sTipoDocumentoAceptador = string.Empty;
                            string sDigitoVerificacionAceptador = string.Empty;
                            string UsuarioSAPActual = string.Empty;
                            string sGetauthorizer = string.Empty;

                            UsuarioSAPActual = Convert.ToString(_oCompany.UserSignature);

                            #endregion

                            #region Valida si esta configurado el usuario                   

                            SAPbobsCOM.Recordset oGetauthorizer = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                            sGetauthorizer = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "Getauthorizer");

                            sGetauthorizer = sGetauthorizer.Replace("%UserId%", UsuarioSAPActual);

                            oGetauthorizer.DoQuery(sGetauthorizer);

                            #endregion

                            if (oGetauthorizer.RecordCount > 0)
                            {
                                sNombreAceptador = Convert.ToString(oGetauthorizer.Fields.Item("Nombre").Value.ToString());
                                sApellidoAceptador = Convert.ToString(oGetauthorizer.Fields.Item("Apellido").Value.ToString());
                                sCargoAceptador = Convert.ToString(oGetauthorizer.Fields.Item("Cargo").Value.ToString());
                                sDepartamentoAceptador = Convert.ToString(oGetauthorizer.Fields.Item("Departamento").Value.ToString());
                                sNITAceptador = Convert.ToString(oGetauthorizer.Fields.Item("NIT").Value.ToString());
                                sTipoDocumentoAceptador = Convert.ToString(oGetauthorizer.Fields.Item("TipoDocumento").Value.ToString());

                                #region Valida campos obligatorios 

                                if (string.IsNullOrEmpty(sNombreAceptador))
                                {
                                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10028", _sboapp, null));
                                }
                                else if (string.IsNullOrEmpty(sApellidoAceptador))
                                {
                                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER100289", _sboapp, null));
                                }
                                else if (string.IsNullOrEmpty(sCargoAceptador))
                                {
                                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10029", _sboapp, null));
                                }
                                else if (string.IsNullOrEmpty(sDepartamentoAceptador))
                                {
                                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10030", _sboapp, null));
                                }
                                else if (string.IsNullOrEmpty(sNITAceptador))
                                {
                                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10031", _sboapp, null));
                                }
                                else if (string.IsNullOrEmpty(sTipoDocumentoAceptador))
                                {
                                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10032", _sboapp, null));
                                }
                                else
                                {

                                    #region Consulta URL

                                    string sGetModo = null;
                                    string sURLRecepcion = null;
                                    string sModo = null;
                                    string sRutaXML = null;
                                    string sRutaPDF = null;
                                    string sProtocoloComunicacion = null;
                                    string sTokenEmpresa = null;
                                    string sTokenPassword = null;
                                    string sGetDV = string.Empty;

                                    SAPbobsCOM.Recordset oConsultarGetModo = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                    sGetModo = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetModoandURL");

                                    sGetModo = sGetModo.Replace("%Estado%", "\"U_BO_Status\" = 'Y'").Replace("%DocEntry%", " ");

                                    oConsultarGetModo.DoQuery(sGetModo);

                                    sURLRecepcion = Convert.ToString(oConsultarGetModo.Fields.Item("URLTFHKA").Value.ToString()) + "/ReceptorWS.svc?wsdl";
                                    sModo = Convert.ToString(oConsultarGetModo.Fields.Item("Modo").Value.ToString());

                                    sProtocoloComunicacion = Convert.ToString(oConsultarGetModo.Fields.Item("ProtocoloComunicacion").Value.ToString());
                                    sTokenEmpresa = Convert.ToString(oConsultarGetModo.Fields.Item("TokenEmpresa").Value.ToString());
                                    sTokenPassword = Convert.ToString(oConsultarGetModo.Fields.Item("TokenPassword").Value.ToString());
                                    sRutaXML = Convert.ToString(oConsultarGetModo.Fields.Item("RutaXML").Value.ToString());
                                    sRutaPDF = Convert.ToString(oConsultarGetModo.Fields.Item("RutaPDF").Value.ToString());

                                    DllFunciones.liberarObjetos(oConsultarGetModo);

                                    #endregion

                                    #region Instanciacion parametros TFHKA

                                    if (sProtocoloComunicacion == "HTTP")
                                    {
                                        BasicHttpBinding port = new BasicHttpBinding();
                                    }
                                    else if (sProtocoloComunicacion == "HTTPS")
                                    {
                                        BasicHttpsBinding port = new BasicHttpsBinding();
                                    }

                                    port.MaxBufferPoolSize = Int32.MaxValue;
                                    port.MaxBufferSize = Int32.MaxValue;
                                    port.MaxReceivedMessageSize = Int32.MaxValue;
                                    port.ReaderQuotas.MaxStringContentLength = Int32.MaxValue;
                                    port.SendTimeout = TimeSpan.FromMinutes(2);
                                    port.ReceiveTimeout = TimeSpan.FromMinutes(2);

                                    if (sProtocoloComunicacion == "HTTPS")
                                    {
                                        port.Security.Mode = BasicHttpSecurityMode.Transport;
                                    }

                                    EndpointAddress endPointEmision = new EndpointAddress(sURLRecepcion); //URL

                                    Recepcion21WS.ReceptorWSClient serviceClienTFHKAReception;
                                    serviceClienTFHKAReception = new Recepcion21WS.ReceptorWSClient(port, endPointEmision);

                                    #endregion

                                    #region Parametros generales Reporte

                                    ReceptorCambioEstatusRequest ParametrosCambioEstatus = new ReceptorCambioEstatusRequest();

                                    ParametrosCambioEstatus.tokenEmpresa = sTokenEmpresa;
                                    ParametrosCambioEstatus.tokenPassword = sTokenPassword;

                                    ParametrosCambioEstatus.identificadorEmisor = sIdentificacionEmisor;
                                    ParametrosCambioEstatus.tipoIdentificacionemisor = sTipoIdentificacionEmisor;
                                    ParametrosCambioEstatus.numeroDocumento = sNumeroDocumentoFacturaProveedor;

                                    ReceptorCambioEstatusRequest.EjecutadoPorRequest UsuarioAceptador = new ReceptorCambioEstatusRequest.EjecutadoPorRequest();

                                    UsuarioAceptador.Nombre = sNombreAceptador;
                                    UsuarioAceptador.Apellido = sApellidoAceptador;
                                    UsuarioAceptador.Cargo = sCargoAceptador;
                                    UsuarioAceptador.Departamento = sDepartamentoAceptador;

                                    ParametrosCambioEstatus.EjecutadoPor = UsuarioAceptador;

                                    ReceptorCambioEstatusRequest.EjecutadoPorRequest.IdentificacionRequest NITAceptador = new ReceptorCambioEstatusRequest.EjecutadoPorRequest.IdentificacionRequest();

                                    NITAceptador.NumeroIdentificacion = sNITAceptador;
                                    NITAceptador.TipoIdentificacion = sTipoDocumentoAceptador;

                                    #region Consulta Digito Verificacion                   

                                    SAPbobsCOM.Recordset oGetDV = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                    sGetDV = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetDV");

                                    sGetDV = sGetDV.Replace("%NIT%", sNITAceptador);

                                    oGetDV.DoQuery(sGetDV);

                                    if (oGetDV.RecordCount > 0)
                                    {
                                        sDigitoVerificacionAceptador = Convert.ToString(oGetDV.Fields.Item("DV").Value.ToString());
                                    }

                                    #endregion

                                    NITAceptador.Dv = sDigitoVerificacionAceptador;

                                    ParametrosCambioEstatus.EjecutadoPor.Identificacion = NITAceptador;

                                    ParametrosCambioEstatus.status = sCodigoEstadoDIAN;
                                    ParametrosCambioEstatus.codigoRechazo = "02";

                                    #endregion

                                    #region Cambia estado en DIAN                            

                                    Recepcion21WS.ResponseGeneral wsCambiarEstado;

                                    wsCambiarEstado = serviceClienTFHKAReception.CambioEstatus(ParametrosCambioEstatus);

                                    if (wsCambiarEstado.codigo == 200)
                                    {
                                        #region Actualiza estado documento en SAP

                                        SAPbobsCOM.Recordset oUpdateStatusDocument = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                        string sUpdateStatusDocument = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "PostUpdateStatusDocument");

                                        sUpdateStatusDocument = sUpdateStatusDocument.Replace("%NumeroFactura%", sNumeroDocumentoFacturaProveedor).Replace("%NumeroIdentificacion%", sIdentificacionEmisor).Replace("%CodigoEventoDIANPT%", sCodigoEstadoDIAN);

                                        oUpdateStatusDocument.DoQuery(sUpdateStatusDocument);

                                        DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10034", _sboapp, null));

                                        #endregion
                                    }
                                    else
                                    {
                                        DllFunciones.sendMessageBox(_sboapp, "RE007 - No se pudo cambiar el estado - " + wsCambiarEstado.mensaje.ToString());
                                    }

                                    #endregion

                                }

                                #endregion

                            }
                        }
                    }
                }

                #endregion
            }

        }

        private void ChagueDocumentStatusBulk_TFHKA(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Form _oFormVisorRepcecion, ItemEvent pVal, string _ColUID)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            #region Valida si esta configurado el usuario     
              
            string sGetauthorizer = string.Empty;
            string UsuarioSAPActual = string.Empty;
            string sCodigoEstadoDIAN = string.Empty;

            UsuarioSAPActual = Convert.ToString(_oCompany.UserSignature);

            SAPbobsCOM.Recordset oGetauthorizer = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            sGetauthorizer = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "Getauthorizer");

            sGetauthorizer = sGetauthorizer.Replace("%UserId%", UsuarioSAPActual);

            oGetauthorizer.DoQuery(sGetauthorizer);

            #endregion

            if (oGetauthorizer.RecordCount > 0)
            {
                #region Validacion campos obligatorios y genera evento

                string sNombreAceptador = Convert.ToString(oGetauthorizer.Fields.Item("Nombre").Value.ToString());
                string sApellidoAceptador = Convert.ToString(oGetauthorizer.Fields.Item("Apellido").Value.ToString());
                string sCargoAceptador = Convert.ToString(oGetauthorizer.Fields.Item("Cargo").Value.ToString());
                string sDepartamentoAceptador = Convert.ToString(oGetauthorizer.Fields.Item("Departamento").Value.ToString());
                string sNITAceptador = Convert.ToString(oGetauthorizer.Fields.Item("NIT").Value.ToString());
                string sTipoDocumentoAceptador = Convert.ToString(oGetauthorizer.Fields.Item("TipoDocumento").Value.ToString());

                if (string.IsNullOrEmpty(sNombreAceptador))
                {
                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10028", _sboapp, null));
                }
                else if (string.IsNullOrEmpty(sApellidoAceptador))
                {
                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10029", _sboapp, null));
                }
                else if (string.IsNullOrEmpty(sCargoAceptador))
                {
                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10030", _sboapp, null));
                }
                else if (string.IsNullOrEmpty(sDepartamentoAceptador))
                {
                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10031", _sboapp, null));
                }
                else if (string.IsNullOrEmpty(sNITAceptador))
                {
                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10032", _sboapp, null));
                }
                else if (string.IsNullOrEmpty(sTipoDocumentoAceptador))
                {
                    DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10033", _sboapp, null));
                }
                else
                {
                    #region Genera Eventos Masivamente 

                    #region Variables y Objetos

                    SAPbouiCOM.ButtonCombo cbEvMa = (ButtonCombo)_oFormVisorRepcecion.Items.Item("cbEvMa").Specific;
                    SAPbouiCOM.Matrix oMatrixOPCH = (Matrix)_oFormVisorRepcecion.Items.Item("MtxOPCH").Specific;
                    SAPbouiCOM.ComboBox cboStado = (ComboBox)_oFormVisorRepcecion.Items.Item("cboStado").Specific;

                    var EstadosPermitidosAcuse = new[] { "00", "01", "13" };
                    var EstadosPermitidosRecibo = new[] { "10" };
                    var EstadosPermitidosAceptacion = new[] { "12" };

                    #endregion

                    int selectedCount = 0;

                    DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Warning, MessageSystemAddOn("FER10037", _sboapp, null));

                    #region Valida lineas marcadas a procesar

                    for (int i = 1; i <= oMatrixOPCH.RowCount; i++)
                    {
                        SAPbouiCOM.CheckBox oCheckBox = (SAPbouiCOM.CheckBox)oMatrixOPCH.Columns.Item("Col_28").Cells.Item(i).Specific;

                        if (oCheckBox.Checked)
                        {
                            selectedCount++;
                        }
                    }

                    #endregion

                    if (cbEvMa.Selected == null)
                    {
                        DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10038", _sboapp, null));
                    }
                    else if (selectedCount == 0)
                    {
                        DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10039", _sboapp, null));
                    }
                    else if (string.IsNullOrEmpty(cboStado.Value) || cboStado.Value == "-")
                    {
                        DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10040", _sboapp, null));
                    }
                    else
                    {
                        int continuar = DllFunciones.sendMessageBoxY_N(_sboapp, $"Se generara eventos a {selectedCount.ToString()} documentos, ¿ Desea Continuar ?.");
                        bool EjecutaProceso = false;

                        if (continuar == 1)
                        {
                            #region Validacion de reglas

                            if (cbEvMa.Selected.Value == "10")
                            {
                                if (EstadosPermitidosAcuse.Contains(cboStado.Value))
                                {
                                    EjecutaProceso = true;
                                    sCodigoEstadoDIAN = "10";
                                }
                                else
                                {
                                    DllFunciones.sendMessageBox(_sboapp, "No se puede generar el evento Acuse de recibo, los documentos seleccionados no estan en estado Cargado, Precargado o Entregado, por favor validar.  ");
                                }
                            }
                            else if (cbEvMa.Selected.Value == "12")
                            {
                                if (EstadosPermitidosRecibo.Contains(cboStado.Value))
                                {
                                    EjecutaProceso = true;
                                    sCodigoEstadoDIAN = "12";
                                }
                                else
                                {
                                    DllFunciones.sendMessageBox(_sboapp, "No se puede generar el Recibo del bien y/o prestacion del servicio, los documentos seleccionados no estan en estado Acuse de recibo (DIAN), por favor validar.  ");
                                }
                            }
                            else if (cbEvMa.Selected.Value == "02")
                            {
                                if (EstadosPermitidosAceptacion.Contains(cboStado.Value))
                                {
                                    EjecutaProceso = true;
                                    sCodigoEstadoDIAN = "02";
                                }
                                else
                                {
                                    DllFunciones.sendMessageBox(_sboapp, "No se puede generar La Aceptacion Expresa de la DIAN, los documentos seleccionados no estan en estado Recibo del bien y/o prestacion del servicio, por favor validar.");
                                }
                            }

                            #endregion

                            if (EjecutaProceso)
                            {
                                #region Variable y objetos

                                string sGetModo = null;
                                string sURLRecepcion = null;
                                string sModo = null;
                                string sRutaXML = null;
                                string sRutaPDF = null;
                                string sProtocoloComunicacion = null;
                                string sTokenEmpresa = null;
                                string sTokenPassword = null;
                                string sGetDV = string.Empty;

                                #endregion

                                #region Consulta URL

                                SAPbobsCOM.Recordset oConsultarGetModo = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                sGetModo = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetModoandURL");

                                sGetModo = sGetModo.Replace("%Estado%", "\"U_BO_Status\" = 'Y'").Replace("%DocEntry%", " ");

                                oConsultarGetModo.DoQuery(sGetModo);

                                sURLRecepcion = Convert.ToString(oConsultarGetModo.Fields.Item("URLTFHKA").Value.ToString()) + "/ReceptorWS.svc?wsdl";
                                sModo = Convert.ToString(oConsultarGetModo.Fields.Item("Modo").Value.ToString());

                                sProtocoloComunicacion = Convert.ToString(oConsultarGetModo.Fields.Item("ProtocoloComunicacion").Value.ToString());
                                sTokenEmpresa = Convert.ToString(oConsultarGetModo.Fields.Item("TokenEmpresa").Value.ToString());
                                sTokenPassword = Convert.ToString(oConsultarGetModo.Fields.Item("TokenPassword").Value.ToString());
                                sRutaXML = Convert.ToString(oConsultarGetModo.Fields.Item("RutaXML").Value.ToString());
                                sRutaPDF = Convert.ToString(oConsultarGetModo.Fields.Item("RutaPDF").Value.ToString());

                                DllFunciones.liberarObjetos(oConsultarGetModo);

                                #endregion

                                #region Instanciacion parametros TFHKA

                                if (sProtocoloComunicacion == "HTTP")
                                {
                                    BasicHttpBinding port = new BasicHttpBinding();
                                }
                                else if (sProtocoloComunicacion == "HTTPS")
                                {
                                    BasicHttpsBinding port = new BasicHttpsBinding();
                                }

                                port.MaxBufferPoolSize = Int32.MaxValue;
                                port.MaxBufferSize = Int32.MaxValue;
                                port.MaxReceivedMessageSize = Int32.MaxValue;
                                port.ReaderQuotas.MaxStringContentLength = Int32.MaxValue;
                                port.SendTimeout = TimeSpan.FromMinutes(2);
                                port.ReceiveTimeout = TimeSpan.FromMinutes(2);

                                if (sProtocoloComunicacion == "HTTPS")
                                {
                                    port.Security.Mode = BasicHttpSecurityMode.Transport;
                                }

                                EndpointAddress endPointEmision = new EndpointAddress(sURLRecepcion);

                                Recepcion21WS.ReceptorWSClient serviceClienTFHKAReception;
                                serviceClienTFHKAReception = new Recepcion21WS.ReceptorWSClient(port, endPointEmision);

                                #endregion

                                DllFunciones.ProgressBar(_oCompany, _sboapp, 3, 1, MessageSystemAddOn("FER10041", _sboapp, null));
                                DllFunciones.ProgressBar(_oCompany, _sboapp, 3, 1, MessageSystemAddOn("FER10041", _sboapp, null));

                                for (int i = 1; i <= oMatrixOPCH.RowCount; i++)
                                {
                                    SAPbouiCOM.CheckBox oCheckBox = (SAPbouiCOM.CheckBox)oMatrixOPCH.Columns.Item("Col_28").Cells.Item(i).Specific;

                                    if (oCheckBox.Checked)
                                    {
                                        selectedCount++;

                                        #region Variables y Objetos

                                        string sCondicionPago = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_24").Cells.Item(i).Specific)).Value;
                                        string sNumeroDocumentoFacturaProveedor = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_1").Cells.Item(i).Specific)).Value;
                                        string sIdentificacionEmisor = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_2").Cells.Item(i).Specific)).Value;
                                        string sTipoIdentificacionEmisor = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_14").Cells.Item(i).Specific)).Value;
                                        string sNombreProveedor = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_3").Cells.Item(i).Specific)).Value;
                                        string sEstadoDIAN = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_0").Cells.Item(i).Specific)).Value;
                                        string sDigitoVerificacionAceptador = string.Empty;

                                        #endregion

                                        #region Procesa evento DIAN

                                        #region Parametros generales Reporte

                                        ReceptorCambioEstatusRequest ParametrosCambioEstatus = new ReceptorCambioEstatusRequest();

                                        ParametrosCambioEstatus.tokenEmpresa = sTokenEmpresa;
                                        ParametrosCambioEstatus.tokenPassword = sTokenPassword;

                                        ParametrosCambioEstatus.identificadorEmisor = sIdentificacionEmisor;
                                        ParametrosCambioEstatus.tipoIdentificacionemisor = sTipoIdentificacionEmisor;
                                        ParametrosCambioEstatus.numeroDocumento = sNumeroDocumentoFacturaProveedor;

                                        ReceptorCambioEstatusRequest.EjecutadoPorRequest UsuarioAceptador = new ReceptorCambioEstatusRequest.EjecutadoPorRequest();

                                        UsuarioAceptador.Nombre = sNombreAceptador;
                                        UsuarioAceptador.Apellido = sApellidoAceptador;
                                        UsuarioAceptador.Cargo = sCargoAceptador;
                                        UsuarioAceptador.Departamento = sDepartamentoAceptador;

                                        ParametrosCambioEstatus.EjecutadoPor = UsuarioAceptador;

                                        ReceptorCambioEstatusRequest.EjecutadoPorRequest.IdentificacionRequest NITAceptador = new ReceptorCambioEstatusRequest.EjecutadoPorRequest.IdentificacionRequest();

                                        NITAceptador.NumeroIdentificacion = sNITAceptador;
                                        NITAceptador.TipoIdentificacion = sTipoDocumentoAceptador;

                                        #region Consulta Digito Verificacion                   

                                        SAPbobsCOM.Recordset oGetDV = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                        sGetDV = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetDV");

                                        sGetDV = sGetDV.Replace("%NIT%", sNITAceptador);

                                        oGetDV.DoQuery(sGetDV);

                                        if (oGetDV.RecordCount > 0)
                                        {
                                            sDigitoVerificacionAceptador = Convert.ToString(oGetDV.Fields.Item("DV").Value.ToString());
                                        }

                                        #endregion

                                        NITAceptador.Dv = sDigitoVerificacionAceptador;

                                        ParametrosCambioEstatus.EjecutadoPor.Identificacion = NITAceptador;

                                        ParametrosCambioEstatus.status = sCodigoEstadoDIAN;
                                        ParametrosCambioEstatus.codigoRechazo = "02";

                                        #endregion

                                        #region Cambia estado en DIAN                            

                                        Recepcion21WS.ResponseGeneral wsCambiarEstado;

                                        wsCambiarEstado = serviceClienTFHKAReception.CambioEstatus(ParametrosCambioEstatus);

                                        if (wsCambiarEstado.codigo == 200)
                                        {
                                            #region Actualiza estado documento en SAP

                                            SAPbobsCOM.Recordset oUpdateStatusDocument = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                            string sUpdateStatusDocument = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "PostUpdateStatusDocument");

                                            sUpdateStatusDocument = sUpdateStatusDocument.Replace("%NumeroFactura%", sNumeroDocumentoFacturaProveedor).Replace("%NumeroIdentificacion%", sIdentificacionEmisor).Replace("%CodigoEventoDIANPT%", sCodigoEstadoDIAN);

                                            oUpdateStatusDocument.DoQuery(sUpdateStatusDocument);

                                            #endregion
                                        }


                                        #endregion
                                        
                                        #endregion

                                    }
                                }

                                DllFunciones.ProgressBar(_oCompany, _sboapp, 3, 1, MessageSystemAddOn("FER10041", _sboapp, null));

                                DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, MessageSystemAddOn("FER10042", _sboapp, null));
                            }
                        }
                    }

                    #endregion
                }

                #endregion                
            }
            else
            {
                DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10035",_sboapp, null));
            }            
        }

        private void ExportToExcel_TFHKA(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Form _oFormVisorRepcecion)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {
                DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Warning, MessageSystemAddOn("FER10043", _sboapp, null));

                string sNombrePlantillaXLS = "InformeFacturasProveedor.xlsx";
                string sPathXLS = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\BOElectronicReception\\Excel\\" + sNombrePlantillaXLS;

                // 2. Cargar plantilla
                ExcelHX.Application xlApp = new ExcelHX.Application();
                ExcelHX.Workbook xlWorkBook = xlApp.Workbooks.Open(sPathXLS);
                ExcelHX.Worksheet xlSheet = (ExcelHX.Worksheet)xlWorkBook.Sheets[1];

                // 3. Escribir datos a partir de fila 4 (ajusta según tu plantilla)
                int startRow = 4;
                int currentRow = startRow;

                SAPbouiCOM.ButtonCombo cbEvMa = (ButtonCombo)_oFormVisorRepcecion.Items.Item("cbEvMa").Specific;
                SAPbouiCOM.Matrix oMatrixOPCH = (Matrix)_oFormVisorRepcecion.Items.Item("MtxOPCH").Specific;
                SAPbouiCOM.ComboBox cboStado = (ComboBox)_oFormVisorRepcecion.Items.Item("cboStado").Specific;

                for (int i = 1; i <= oMatrixOPCH.RowCount; i++)
                {

                    xlSheet.Cells[currentRow, 1] = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_0").Cells.Item(i).Specific)).Value;
                    xlSheet.Cells[currentRow, 2] = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_11").Cells.Item(i).Specific)).Value;
                    xlSheet.Cells[currentRow, 3] = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_29").Cells.Item(i).Specific)).Value;
                    xlSheet.Cells[currentRow, 4] = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_1").Cells.Item(i).Specific)).Value;
                    xlSheet.Cells[currentRow, 5] = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_18").Cells.Item(i).Specific)).Value;
                    xlSheet.Cells[currentRow, 6] = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_10").Cells.Item(i).Specific)).Value;
                    xlSheet.Cells[currentRow, 7] = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_2").Cells.Item(i).Specific)).Value;
                    xlSheet.Cells[currentRow, 8] = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_3").Cells.Item(i).Specific)).Value;
                    xlSheet.Cells[currentRow, 9] = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_6").Cells.Item(i).Specific)).Value;
                    xlSheet.Cells[currentRow, 10] = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_17").Cells.Item(i).Specific)).Value;
                    xlSheet.Cells[currentRow, 11] = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_8").Cells.Item(i).Specific)).Value;
                    
                    currentRow++;
                }

                // Guardar como nuevo archivo (para no sobrescribir plantilla)

                // 1. Obtener ruta del escritorio
                string escritorioPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                // 2. Crear el nombre del archivo con la fecha actual
                string fecha = DateTime.Now.ToString("yyyy-MM-dd"); // o "yyyyMMdd" si prefieres sin guiones
                string nombreArchivo = $"InformeFacturasProveedor_{fecha}.xlsx";

                // 3. Combinar ruta y nombre
                string nuevoPath = System.IO.Path.Combine(escritorioPath, nombreArchivo);

                xlWorkBook.SaveAs(nuevoPath);
                xlWorkBook.Close(false);
                xlApp.Quit();

                // 5. Liberar recursos
                Marshal.ReleaseComObject(xlSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);

                string[] arrNombreArchivo = { nombreArchivo };

                DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10044", _sboapp, arrNombreArchivo));


            }
            catch (Exception)
            {

                throw;
            }

        }

        #endregion

        #region Proveedor Tecnologico Factura By Estela 

        private void DescargaDocumentosFBE(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Form _oFormVisorRepcecion)
        {
            try
            {
                int iReprocesar = DllFunciones.sendMessageBoxY_N(_sboapp, MessageSystemAddOn("FER10001", _sboapp, null));

                if (iReprocesar == 1)
                {
                    #region Consulta URL

                    string sGetModo = null;                                        
                    string sAPIFBE = null;
                    string sxWho = null;
                    string sUserFBE = null;
                    string sPassFBE = null;
                    string sTenantId = null;

                    SAPbobsCOM.Recordset oConsultarGetModo = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    SAPbobsCOM.Recordset oSyncDocsRecep = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    sGetModo = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetModoandURLFBE");

                    sGetModo = sGetModo.Replace("WHERE %Estado% %DocEntry%", "");

                    oConsultarGetModo.DoQuery(sGetModo);

                    sAPIFBE = Convert.ToString(oConsultarGetModo.Fields.Item("APIFBE").Value.ToString());
                    sxWho = Convert.ToString(oConsultarGetModo.Fields.Item("xWho").Value.ToString());
                    sUserFBE = Convert.ToString(oConsultarGetModo.Fields.Item("UserFBE").Value.ToString());
                    sPassFBE = Convert.ToString(oConsultarGetModo.Fields.Item("PassFBE").Value.ToString());
                    sTenantId = Convert.ToString(oConsultarGetModo.Fields.Item("TenatIdFBE").Value.ToString());

                    DllFunciones.liberarObjetos(oConsultarGetModo);

                    #endregion

                    #region Variables y Objetos

                    string sPathImages = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\BOElectronicReception\\Images\\";

                    #endregion

                    var AccesTokenFBE = FBE_GetAuthorizationAccess(_sboapp, _oCompany, sAPIFBE, sUserFBE, sPassFBE, sTenantId, sxWho).GetAwaiter().GetResult();

                    if (AccesTokenFBE == "401")
                    {
                        DllFunciones.sendMessageBox(_sboapp, MessageSystemAddOn("FER10011", _sboapp, null));
                    }
                    else
                    {
                        var dtResponsive = SearchDocumentsReceipt(_sboapp, _oCompany, sAPIFBE, sxWho, AccesTokenFBE).GetAwaiter().GetResult();



                    }
                                        

                    #region Carga Infortmacion en la Matrix

                    //#region Variables y Objetos

                    //string sPath;
                    //string sInvoices = null;
                    //string sCreditMemo = null;
                    //string sDebitMemo = null;

                    //int CantidadRegistos = 0;

                    //#endregion

                    //sPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

                    //#region Consulta de documentos facturas, notas debito y notas credito a mostrar en matrix

                    //SAPbouiCOM.Matrix oMatrixInvoice = (Matrix)_oFormVisorRepcecion.Items.Item("MtxOPCH").Specific;
                    //SAPbouiCOM.Matrix oMatrixCreditMemo = (Matrix)_oFormVisorRepcecion.Items.Item("MtxORPC").Specific;
                    //SAPbouiCOM.Matrix oMatrixDebitMemo = (Matrix)_oFormVisorRepcecion.Items.Item("MtxOPCHD").Specific;
                    //SAPbouiCOM.ComboBox oEstado = (SAPbouiCOM.ComboBox)_oFormVisorRepcecion.Items.Item("cboStado").Specific;

                    //SAPbobsCOM.Recordset oRecorsetInvoices = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    //SAPbobsCOM.Recordset oRecorsetCreditMemo = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    //SAPbobsCOM.Recordset oRecorsetDebitMemo = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    //SAPbouiCOM.DataTable oTableInvoices = _oFormVisorRepcecion.DataSources.DataTables.Item("DT_Invoices");
                    //SAPbouiCOM.DataTable oTableCreditMemo = _oFormVisorRepcecion.DataSources.DataTables.Item("DT_CreditMemo");
                    //SAPbouiCOM.DataTable oTableDebitMemo = _oFormVisorRepcecion.DataSources.DataTables.Item("DT_DebitMemo");

                    //sInvoices = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetInvoices");
                    //sInvoices = sInvoices.Replace("%PathImages%", sPathImages).Replace("%FI%", "20220108").Replace("%FF%", "20251231");

                    //sCreditMemo = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetCreditMemo");
                    //sInvoices = sInvoices.Replace("%PathImages%", sPathImages).Replace("%FI%", "20220108").Replace("%FF%", "20251231");

                    //sDebitMemo = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetDebitMemo");
                    //sInvoices = sInvoices.Replace("%PathImages%", sPathImages).Replace("%FI%", "20220108").Replace("%FF%", "20251231");

                    //if (oEstado.Value == "-" || oEstado.Value == "")
                    //{
                    //    sInvoices = sInvoices.Replace("%Estado%", "");
                    //}
                    //else
                    //{
                    //    sInvoices = sInvoices.Replace("%Estado%", "AND \"U_BOTSDC\" = '" + oEstado.Value + "' ");
                    //}


                    //oRecorsetInvoices.DoQuery(sInvoices);
                    //oRecorsetCreditMemo.DoQuery(sCreditMemo);
                    //oRecorsetDebitMemo.DoQuery(sDebitMemo);

                    //oTableInvoices.ExecuteQuery(sInvoices);
                    //oTableCreditMemo.ExecuteQuery(sCreditMemo);
                    //oTableDebitMemo.ExecuteQuery(sDebitMemo);

                    //#endregion

                    //CantidadRegistos = oRecorsetInvoices.RecordCount + oRecorsetCreditMemo.RecordCount + oRecorsetDebitMemo.RecordCount;

                    //if (CantidadRegistos != 0)
                    //{
                    //    #region Carga datos Matrix Facturas

                    //    if (oRecorsetInvoices.RecordCount > 0)
                    //    {
                    //        oMatrixInvoice.Clear();

                    //        oMatrixInvoice.Columns.Item("#").DataBind.Bind("DT_Invoices", "#");
                    //        oMatrixInvoice.Columns.Item("Col_0").DataBind.Bind("DT_Invoices", "Estado");
                    //        oMatrixInvoice.Columns.Item("Col_11").DataBind.Bind("DT_Invoices", "Estado_SAP");
                    //        oMatrixInvoice.Columns.Item("Col_9").DataBind.Bind("DT_Invoices", "DocEntry");
                    //        oMatrixInvoice.Columns.Item("Col_1").DataBind.Bind("DT_Invoices", "Num_Fac_Pro");
                    //        oMatrixInvoice.Columns.Item("Col_16").DataBind.Bind("DT_Invoices", "Serie_Numeracion");
                    //        oMatrixInvoice.Columns.Item("Col_2").DataBind.Bind("DT_Invoices", "Codigo_cliente");
                    //        oMatrixInvoice.Columns.Item("Col_3").DataBind.Bind("DT_Invoices", "Nombre_del_Cliente");
                    //        oMatrixInvoice.Columns.Item("Col_4").DataBind.Bind("DT_Invoices", "Fecha_Documento");
                    //        oMatrixInvoice.Columns.Item("Col_5").DataBind.Bind("DT_Invoices", "Fecha_vencimiento");
                    //        oMatrixInvoice.Columns.Item("Col_24").DataBind.Bind("DT_Invoices", "Condicion_de_Pago");
                    //        oMatrixInvoice.Columns.Item("Col_6").DataBind.Bind("DT_Invoices", "Total_documento");
                    //        oMatrixInvoice.Columns.Item("Col_17").DataBind.Bind("DT_Invoices", "CUFE");
                    //        oMatrixInvoice.Columns.Item("Col_18").DataBind.Bind("DT_Invoices", "Fecha_emision");
                    //        oMatrixInvoice.Columns.Item("Col_10").DataBind.Bind("DT_Invoices", "Hora_emision");
                    //        oMatrixInvoice.Columns.Item("Col_8").DataBind.Bind("DT_Invoices", "Fecha_recepcion");
                    //        oMatrixInvoice.Columns.Item("Col_19").Visible = false;
                    //        oMatrixInvoice.Columns.Item("Col_23").Visible = false;
                    //        oMatrixInvoice.Columns.Item("Col_20").DataBind.Bind("DT_Invoices", "DescargaXML");
                    //        oMatrixInvoice.Columns.Item("Col_12").DataBind.Bind("DT_Invoices", "RutaXML");
                    //        oMatrixInvoice.Columns.Item("Col_12").Visible = false;
                    //        oMatrixInvoice.Columns.Item("Col_21").DataBind.Bind("DT_Invoices", "DescargaPDF");
                    //        oMatrixInvoice.Columns.Item("Col_13").DataBind.Bind("DT_Invoices", "RutaPDF");
                    //        oMatrixInvoice.Columns.Item("Col_13").Visible = false;
                    //        oMatrixInvoice.Columns.Item("Col_22").DataBind.Bind("DT_Invoices", "ImageAceptar");
                    //        oMatrixInvoice.Columns.Item("Col_7").DataBind.Bind("DT_Invoices", "ImageCancelar");
                    //        oMatrixInvoice.Columns.Item("Col_14").DataBind.Bind("DT_Invoices", "TipoIdentificacionEmisor");
                    //        oMatrixInvoice.Columns.Item("Col_14").Visible = false;

                    //        oMatrixInvoice.LoadFromDataSource();

                    //        oMatrixInvoice.AutoResizeColumns();

                    //    }

                    //    #endregion

                    //    #region Carga datos Matrix Notas credito

                    //    if (oRecorsetCreditMemo.RecordCount > 0)
                    //    {
                    //        oMatrixCreditMemo.Clear();

                    //        oMatrixCreditMemo.Columns.Item("#").DataBind.Bind("DT_Invoices", "#");
                    //        oMatrixCreditMemo.Columns.Item("Col_0").DataBind.Bind("DT_Invoices", "Estado");
                    //        oMatrixCreditMemo.Columns.Item("Col_9").DataBind.Bind("DT_Invoices", "DocEntry");
                    //        oMatrixCreditMemo.Columns.Item("Col_1").DataBind.Bind("DT_Invoices", "Num_Fac_Pro");
                    //        //oMatrixCreditMemo.Columns.Item("Col_25").DataBind.Bind("DT_Invoices", "Num_Fac_Preeli");
                    //        //oMatrixCreditMemo.Columns.Item("Col_26").DataBind.Bind("DT_Invoices", "Num_Fac_SAP");
                    //        oMatrixCreditMemo.Columns.Item("Col_16").DataBind.Bind("DT_Invoices", "Serie_Numeracion");
                    //        oMatrixCreditMemo.Columns.Item("Col_2").DataBind.Bind("DT_Invoices", "Codigo_cliente");
                    //        oMatrixCreditMemo.Columns.Item("Col_3").DataBind.Bind("DT_Invoices", "Nombre_del_Cliente");
                    //        oMatrixCreditMemo.Columns.Item("Col_4").DataBind.Bind("DT_Invoices", "Fecha_Documento");
                    //        oMatrixCreditMemo.Columns.Item("Col_5").DataBind.Bind("DT_Invoices", "Fecha_vencimiento");
                    //        oMatrixCreditMemo.Columns.Item("Col_24").DataBind.Bind("DT_Invoices", "Condicion_Pago");
                    //        oMatrixCreditMemo.Columns.Item("Col_6").DataBind.Bind("DT_Invoices", "Total_documento");
                    //        oMatrixCreditMemo.Columns.Item("Col_17").DataBind.Bind("DT_Invoices", "CUFE");
                    //        oMatrixCreditMemo.Columns.Item("Col_18").DataBind.Bind("DT_Invoices", "Fecha_emision");
                    //        oMatrixCreditMemo.Columns.Item("Col_10").DataBind.Bind("DT_Invoices", "Hora_emision");
                    //        oMatrixCreditMemo.Columns.Item("Col_8").DataBind.Bind("DT_Invoices", "Fecha_recepcion");

                    //        oMatrixCreditMemo.LoadFromDataSource();

                    //        oMatrixCreditMemo.AutoResizeColumns();

                    //    }
                    //    #endregion

                    //    #region Carga datos Matrix Notas Debito

                    //    if (oRecorsetDebitMemo.RecordCount > 0)
                    //    {
                    //        oMatrixDebitMemo.Clear();

                    //        oMatrixDebitMemo.Columns.Item("#").DataBind.Bind("DT_Invoices", "#");
                    //        oMatrixDebitMemo.Columns.Item("Col_0").DataBind.Bind("DT_Invoices", "Estado");
                    //        oMatrixDebitMemo.Columns.Item("Col_9").DataBind.Bind("DT_Invoices", "DocEntry");
                    //        oMatrixDebitMemo.Columns.Item("Col_1").DataBind.Bind("DT_Invoices", "Num_Fac_Pro");
                    //        //oMatrixDebitMemo.Columns.Item("Col_25").DataBind.Bind("DT_Invoices", "Num_Fac_Preeli");
                    //        //oMatrixDebitMemo.Columns.Item("Col_26").DataBind.Bind("DT_Invoices", "Num_Fac_SAP");
                    //        oMatrixDebitMemo.Columns.Item("Col_16").DataBind.Bind("DT_Invoices", "Serie_Numeracion");
                    //        oMatrixDebitMemo.Columns.Item("Col_2").DataBind.Bind("DT_Invoices", "Codigo_cliente");
                    //        oMatrixDebitMemo.Columns.Item("Col_3").DataBind.Bind("DT_Invoices", "Nombre_del_Cliente");
                    //        oMatrixDebitMemo.Columns.Item("Col_4").DataBind.Bind("DT_Invoices", "Fecha_Documento");
                    //        oMatrixDebitMemo.Columns.Item("Col_5").DataBind.Bind("DT_Invoices", "Fecha_vencimiento");
                    //        oMatrixDebitMemo.Columns.Item("Col_24").DataBind.Bind("DT_Invoices", "Condicion_Pago");
                    //        oMatrixDebitMemo.Columns.Item("Col_6").DataBind.Bind("DT_Invoices", "Total_documento");
                    //        oMatrixDebitMemo.Columns.Item("Col_17").DataBind.Bind("DT_Invoices", "CUFE");
                    //        oMatrixDebitMemo.Columns.Item("Col_18").DataBind.Bind("DT_Invoices", "Fecha_emision");
                    //        oMatrixDebitMemo.Columns.Item("Col_10").DataBind.Bind("DT_Invoices", "Hora_emision");
                    //        oMatrixDebitMemo.Columns.Item("Col_8").DataBind.Bind("DT_Invoices", "Fecha_recepcion");

                    //        oMatrixDebitMemo.LoadFromDataSource();

                    //        oMatrixDebitMemo.AutoResizeColumns();

                    //    }
                    //    #endregion

                    //}
                    //else
                    //{
                    //    DllFunciones.sendMessageBox(_sboapp, "No se encontraron documentos");
                    //}

                    //DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Documentos sicronizados correctamente.");

                    #endregion

                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        private async Task<System.Data.DataTable> SearchDocumentsReceipt(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _oCompany, string sApiFBE, string sX_Who, string sAuthorization)
        {
            #region Realiza validacion de los datos de conexion a Facture

            var APIListarDocuments = sApiFBE + "/InboundDocument/Integrated/Document";

            ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;

            DateTime FechaInicial = DateTime.Now.AddMonths(-3);
            DateTime FechaFinal = DateTime.Now;

            string sFechaInicial = FechaInicial.ToString("yyyy-MM-ddTHH:mm:ss");
            string sFechaFinal = FechaFinal.ToString("yyyy-MM-ddTHH:mm:ss");

            var jsonRequest = DllFunciones.LoadJSON(_sboapp, "BOElectronicReception", "InboundDocumens.json");
            jsonRequest = jsonRequest.Replace("%FI%", sFechaInicial).Replace("%FF%", sFechaFinal);

            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", sAuthorization);
            _httpClient.DefaultRequestHeaders.Add("X-Who", sX_Who);

            var requestObject = JsonConvert.DeserializeObject<InboundDocument>(jsonRequest);

            var requestBody = JsonConvert.SerializeObject(requestObject);
            var content = new StringContent(requestBody, Encoding.UTF8, "application/json");            

            var response = await _httpClient.PostAsync(APIListarDocuments, content);

            if (response.StatusCode == HttpStatusCode.OK)
            {
                response.EnsureSuccessStatusCode();

                var responseBody = await response.Content.ReadAsStringAsync();
                JObject jsonResponse = JObject.Parse(responseBody);

                // Crea la estructura de la DataTable
                System.Data.DataTable DataTableResponse = new System.Data.DataTable();
                DataTableResponse.Columns.Add("FechaFactura");
                DataTableResponse.Columns.Add("NumeroFactura");
                DataTableResponse.Columns.Add("NIT");
                DataTableResponse.Columns.Add("RazonSocial");
                DataTableResponse.Columns.Add("FormaPago");
                DataTableResponse.Columns.Add("UltimaFechaActualizacion");
                DataTableResponse.Columns.Add("Evento");
                DataTableResponse.Columns.Add("FechaRecepcion");
                DataTableResponse.Columns.Add("IVA");
                DataTableResponse.Columns.Add("Monto");
                DataTableResponse.Columns.Add("LDF");
                DataTableResponse.Columns.Add("TipoDocumento");
                DataTableResponse.Columns.Add("CUFE");

                // Recorre los elementos del array "eventItems"
                foreach (var item in jsonResponse["items"])
                {
                    DataRow row = DataTableResponse.NewRow();
                    row["FechaFactura"] = item["FechaFactura"];
                    row["NumeroFactura"] = item["NumeroFactura"];
                    row["NIT"] = item["NIT"];
                    row["RazonSocial"] = item["RazonSocial"];
                    row["FormaPago"] = item["FormaPago"];
                    row["UltimaFechaActualizacion"] = item["UltimaFechaActualizacion"];
                    row["Evento"] = item["Evento"];
                    row["FechaRecepcion"] = item["FechaRecepcion"];
                    row["IVA"] = item["IVA"];
                    row["Monto"] = item["Monto"];
                    row["LDF"] = item["LDF"];
                    row["TipoDocumento"] = item["TipoDocumento"];
                    row["CUFE"] = item["CUFE"];

                    DataTableResponse.Rows.Add(row);
                }

                return DataTableResponse;

            }
            else
            {
                var responseBody = await response.Content.ReadAsStringAsync();

                JObject jsonResponse = JObject.Parse(responseBody);

                string[] comodinSearchDocumentsReceipt = { (string)jsonResponse["eventItems"]?[0]?["shortDescription"] };

                string ResponseMessage = MessageSystemAddOn("FER10012", _sboapp, comodinSearchDocumentsReceipt);

                DllFunciones.sendMessageBox(_sboapp, ResponseMessage);

                return null;
            }

            #endregion
            
        }

        private async Task<string> FBE_GetAuthorizationAccess(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _oCompany, string sApiFBE, string sUserFBE, string sPassFBE, string sTenantId, string sxWho)
        {
            try
            {
                #region Realiza validacion de los datos de conexion a Facture

                var APIAuthorization = sApiFBE + "/Auth/Login";

                ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;

                var jsonContent = DllFunciones.LoadJSON(_sboapp, "eBilling", "Authorization.json");
                jsonContent = jsonContent.Replace("%User%", sUserFBE).Replace("%Password%", sPassFBE).Replace("%TenantId%", sTenantId);

                var requestObject = JsonConvert.DeserializeObject<AuthorizationRequest>(jsonContent);

                var requestBody = JsonConvert.SerializeObject(requestObject);
                var content = new StringContent(requestBody, Encoding.UTF8, "application/json");

                _httpClient.DefaultRequestHeaders.Add("X-Who", sxWho);

                var response = await _httpClient.PostAsync(APIAuthorization, content);

                if (response.StatusCode == HttpStatusCode.OK)
                {
                    response.EnsureSuccessStatusCode();

                    var responseBody = await response.Content.ReadAsStringAsync();
                    JObject jsonResponse = JObject.Parse(responseBody);
                    string sAccessToken = (string)jsonResponse["accessToken"];

                    return sAccessToken;
                }
                else
                {
                    var responseBody = await response.Content.ReadAsStringAsync();

                    JObject jsonResponse = JObject.Parse(responseBody);

                    string ResponseMessage = (string)jsonResponse["statusCode"];

                    return ResponseMessage;
                }

                #endregion

            }
            catch (Exception e)
            {
                //DllFunciones.sendMessageBox(_sboapp, e.ToString());
                return e.ToString();

            }

        }

        public void SincronizacionDocsFBE(SAPbobsCOM.Company _oCompany, SAPbouiCOM.Application sboapp, SAPbobsCOM.Recordset oSyncDocsRecep, System.Data.DataTable DataTableResponse)
        {         
            //string sSyncDocsRecepOriginal = null;
            //string sSyncDocsRecepCopia = null;

            //ParametrosConsultaDocumentosRecepcion.status_code = sCodigoStatusDIAN;

            //var ResponsiveReportReception00 = _serviceClienTFHKAReception.ReporteStatus(ParametrosConsultaDocumentosRecepcion);

            //if (ResponsiveReportReception00.codigo == 200)
            //{
            //    #region Sincronizando documentos con el proveedor tecnologico                

            //    sSyncDocsRecepOriginal = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "SyncDocsRecep");

            //    for (int i = 0; i < ResponsiveReportReception00.documentoselectronicos.Count(); i++)
            //    {

            //        sSyncDocsRecepCopia = null;

            //        sSyncDocsRecepCopia = sSyncDocsRecepOriginal;

            //        sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%ProveedorTechnologico%", "TFHKA").Replace("%ConsecutivoTFHKA%", ResponsiveReportReception00.documentoselectronicos[i].correlativoempresa.ToString()).Replace("%CUFE%", ResponsiveReportReception00.documentoselectronicos[i].cufe.ToString());

            //        #region StatusDIANCodigo

            //        if (string.IsNullOrEmpty(ResponsiveReportReception00.documentoselectronicos[i].estatusDIANcodigo.ToString()))
            //        {
            //            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%StatusDIANCodigo%'", _sCodigoStatusDIAN);
            //        }
            //        else
            //        {
            //            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%StatusDIANCodigo%", ResponsiveReportReception00.documentoselectronicos[i].estatusDIANcodigo.ToString());
            //        }

            //        #endregion

            //        #region StatusDIANDescripcion

            //        if (string.IsNullOrEmpty(ResponsiveReportReception00.documentoselectronicos[i].estatusDIANdescripcion.ToString()))
            //        {
            //            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%StatusDIANDescripcion%'", "NULL");
            //        }
            //        else
            //        {
            //            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%StatusDIANDescripcion%", ResponsiveReportReception00.documentoselectronicos[i].estatusDIANdescripcion.ToString());
            //        }

            //        #endregion

            //        #region StatusDIANFecha

            //        if (string.IsNullOrEmpty(ResponsiveReportReception00.documentoselectronicos[i].estatusDIANfecha.ToString()))
            //        {
            //            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%StatusDIANFecha%'", "NULL");
            //        }
            //        else
            //        {
            //            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%StatusDIANFecha%", ResponsiveReportReception00.documentoselectronicos[i].estatusDIANfecha.ToString());
            //        }

            //        #endregion

            //        #region FechaEmision

            //        if (string.IsNullOrEmpty(ResponsiveReportReception00.documentoselectronicos[i].fechaemision.ToString()))
            //        {
            //            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%FechaEmision%'", "NULL");
            //        }
            //        else
            //        {
            //            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%FechaEmision%", ResponsiveReportReception00.documentoselectronicos[i].fechaemision.ToString());
            //        }

            //        #endregion

            //        #region FechaRecepcion

            //        if (string.IsNullOrEmpty(ResponsiveReportReception00.documentoselectronicos[i].fecharecepcion.ToString()))
            //        {
            //            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%FechaRecepcion%'", "NULL");
            //        }
            //        else
            //        {
            //            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%FechaRecepcion%", ResponsiveReportReception00.documentoselectronicos[i].fecharecepcion.ToString());
            //        }

            //        #endregion

            //        #region HoraEmision

            //        if (string.IsNullOrEmpty(ResponsiveReportReception00.documentoselectronicos[i].horaemision.ToString()))
            //        {
            //            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%HoraEmision%'", "NULL");
            //        }
            //        else
            //        {
            //            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%HoraEmision%", ResponsiveReportReception00.documentoselectronicos[i].horaemision.ToString());
            //        }

            //        #endregion

            //        #region MontoTotal

            //        if (string.IsNullOrEmpty(ResponsiveReportReception00.documentoselectronicos[i].montototal.ToString()))
            //        {
            //            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%MontoTotal%'", "NULL");
            //        }
            //        else
            //        {
            //            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%MontoTotal%", ResponsiveReportReception00.documentoselectronicos[i].montototal.ToString());
            //        }

            //        #endregion

            //        #region NumeroFactura

            //        if (string.IsNullOrEmpty(ResponsiveReportReception00.documentoselectronicos[i].numerodocumento.ToString()))
            //        {
            //            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%NumeroFactura%'", "NULL");
            //        }
            //        else
            //        {
            //            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%NumeroFactura%", ResponsiveReportReception00.documentoselectronicos[i].numerodocumento.ToString());
            //        }

            //        #endregion

            //        #region NumeroIdentificacion

            //        if (string.IsNullOrEmpty(ResponsiveReportReception00.documentoselectronicos[i].numeroidentificacion.ToString()))
            //        {
            //            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%NumeroIdentificacion%'", "NULL");
            //        }
            //        else
            //        {
            //            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%NumeroIdentificacion%", ResponsiveReportReception00.documentoselectronicos[i].numeroidentificacion.ToString());
            //        }

            //        #endregion

            //        #region RazonSocial

            //        if (string.IsNullOrEmpty(ResponsiveReportReception00.documentoselectronicos[i].razonsocial.ToString()))
            //        {
            //            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%RazonSocial%'", "NULL");
            //        }
            //        else
            //        {
            //            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%RazonSocial%", ResponsiveReportReception00.documentoselectronicos[i].razonsocial.ToString());
            //        }

            //        #endregion

            //        #region tipodocumento

            //        if (string.IsNullOrEmpty(ResponsiveReportReception00.documentoselectronicos[i].tipodocumento.ToString()))
            //        {
            //            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%TipoDocumento%'", "NULL");
            //        }
            //        else
            //        {
            //            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%TipoDocumento%", ResponsiveReportReception00.documentoselectronicos[i].tipodocumento.ToString());
            //        }

            //        #endregion

            //        #region TipoEmisor

            //        if (string.IsNullOrEmpty(ResponsiveReportReception00.documentoselectronicos[i].tipoemisor.ToString()))
            //        {
            //            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%TipoEmisor%'", "NULL");
            //        }
            //        else
            //        {
            //            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%TipoEmisor%", ResponsiveReportReception00.documentoselectronicos[i].tipoemisor.ToString());
            //        }

            //        #endregion

            //        #region TipoIdentidad

            //        if (string.IsNullOrEmpty(ResponsiveReportReception00.documentoselectronicos[i].tipoidentidad.ToString()))
            //        {
            //            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("'%TipoIdentidad%'", "NULL");
            //        }
            //        else
            //        {
            //            sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%TipoIdentidad%", ResponsiveReportReception00.documentoselectronicos[i].tipoidentidad.ToString());
            //        }

            //        #endregion

            //        #region Codigo Estatus DIAN

            //        sSyncDocsRecepCopia = sSyncDocsRecepCopia.Replace("%prmCodEstatusDIAN%", sCodigoStatusDIAN);

            //        #endregion

            //        oSyncDocsRecep.DoQuery(sSyncDocsRecepCopia);

            //    }

            //    #endregion
            //}

        }
        
        #region Clases JSON

        private class InboundDocument
        {
            public string PageNumber { get; set; }
            public string PageSize { get; set; }
            public string DocumentTypeCode { get; set; }
            public string ReceiverStartingDate { get; set; }
            public string ReceiverEndingDate { get; set; }
            public string Lifecycle_Status { get; set; }
            public string FormaPago { get; set; }
            public string LDF { get; set; }
            public string DocumentNumberFull { get; set; }
        }

        private class AuthorizationRequest
        {
            public string u { get; set; }
            public string p { get; set; }
            public string t { get; set; }
        }

        #endregion
        
        #endregion

        private string MessageSystemAddOn(string idMessage, SAPbouiCOM.Application SboAppFacturacionElectronica, string[] sComodinMessage)
        {
            #region Mensajes del sistema

            if (idMessage == "FER10001")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {
                    return idMessage + " - Se descargan los documentos de compras del portal del Proveedor Tecnologico, ¿ Desea continuar ?";
                }
                else
                {
                    return idMessage + " - The purchasing documents are downloaded from the Technology Provider portal. Do you wish to continue?";
                }
            }
            else if (idMessage == "FER10002")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - Sincronizando todos los documentos, Esto puede tomar bastante tiempo, por favor espere...";
                }
                else
                {
                    return idMessage + " - Synchronizing all documents, This may take some time, please wait...";
                }
            }
            else if (idMessage == "FER10003")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - Sincronizando documentos Estado - 00 - Cargado, por favor espere...";
                }
                else
                {
                    return idMessage + " - Synchronizing documents Status - 00 - Loaded, please wait...";
                }
            }
            else if (idMessage == "FER10004")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - Descargando documentos Estado - 01 - Entregado, por favor espere...";
                }
                else
                {
                    return idMessage + " - Downloading documents Status - 01 - Delivered, please wait...";
                }
            }
            else if (idMessage == "FER10005")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - Descargando documentos Estado - 02 - Aceptación expresa (DIAN), por favor espere...";
                }
                else
                {
                    return idMessage + " - Downloading documents Status - 02 - Express acceptance (DIAN), please wait...";
                }
            }
            else if (idMessage == "FER10006")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - Descargando documentos Estado - 02 - Aceptación expresa (DIAN), por favor espere...";
                }
                else
                {
                    return idMessage + " - Downloading documents Status - 04 - Claim (DIAN), please wait...";
                }
            }
            else if (idMessage == "FER10007")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - Descargando documentos Estado - 10 - Acuse de recibo (DIAN), por favor espere...";
                }
                else
                {
                    return idMessage + " - Downloading documents Status - 10 - Acknowledgement of receipt (DIAN), please wait...";
                }
            }
            else if (idMessage == "FER10008")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - Descargando documentos Estado - 11 - Rechazado(DIAN), por favor espere...";
                }
                else
                {
                    return idMessage + " - Downloading documents Status - 11 - Rejected (DIAN), please wait...";
                }
            }
            else if (idMessage == "FER10009")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - Descargando documentos Estado - 12 - Recibo del bien y / o prestación del servicio, por favor espere...";
                }
                else
                {
                    return idMessage + " - Downloading documents Status - 12 - Receipt of the goods and/or provision of the service, please wait...";
                }
            }
            else if (idMessage == "FER10010")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - Descargando documentos Estado - 13 - Precargado, por favor espere...";
                }
                else
                {
                    return idMessage + " - Downloading documents Status - 13 - Preloaded, please wait...";
                }
            }
            else if (idMessage == "FER10011")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - No se ha podido generar el token de acceso para poder sincronizar los documentos, por favor valide sus parametros de conexion ";
                }
                else
                {
                    return idMessage + " - The access token could not be generated to synchronize the documents, please validate your connection parameters";
                }
            }
            else if (idMessage == "FER10012")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - No pudo descargar los documentos del proveedor tecnologico, el motivo es: " + sComodinMessage[0];
                }
                else
                {
                    return idMessage + " - I was unable to download the documents from the technology provider, the reason is: "+ sComodinMessage[0];
                }
            }
            else if (idMessage == "FER10013")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - Se descargan los documentos de compras pendientes, ¿ Desea continuar ? " ;
                }
                else
                {
                    return idMessage + " - Pending purchase documents are downloaded. Do you want to continue? ";
                }
            }
            else if (idMessage == "FER10014")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - Sincronizando documentos " + sComodinMessage[0] + " " + sComodinMessage[1] + " , Esto puede tomar bastante tiempo, por favor espere...";
                }
                else
                {
                    return idMessage + " - Synchronizing documents " + sComodinMessage[0] + " " + sComodinMessage[1] + ", This may take a long time, please wait... ";
                }
            }
            else if (idMessage == "FER10015")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - Sincronizando documentos Estado - 00 - Cargado, por favor espere...";
                }
                else
                {
                    return idMessage + " - Synchronizing documents Status - 00 - Loaded, please wait...";
                }
            }
            else if (idMessage == "FER10016")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - Descargando documentos Estado - 01 - Entregado, por favor espere...";
                }
                else
                {
                    return idMessage + " - Downloading documents Status - 01 - Delivered, please wait...";
                }
            }            
            else if (idMessage == "FER10017")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - Descargando documentos Estado - 02 - Aceptación expresa (DIAN), por favor espere...";
                }
                else
                {
                    return idMessage + " - Downloading documents Status - 02 - Express acceptance (DIAN), please wait...";
                }
            }
            else if (idMessage == "FER10018")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - Descargando documentos Estado - 04 - Reclamo (DIAN), por favor espere...";
                }
                else
                {
                    return idMessage + " - Downloading State documents - 04 - Claim (DIAN), please wait...";
                }
            }
            else if (idMessage == "FER10019")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - Descargando documentos Estado - 10 - Acuse de recibo (DIAN), por favor espere...";
                }
                else
                {
                    return idMessage + " - Downloading documents Status - 10 - Acknowledgment of receipt (DIAN), please wait...";
                }
            }
            else if (idMessage == "FER10020")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - Descargando documentos Estado - 11 - Rechazado (DIAN), por favor espere...";
                }
                else
                {
                    return idMessage + " - Downloading documents Status - 11 - Rejected (DIAN), please wait...";
                }
            }
            else if (idMessage == "FER10021")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - Descargando documentos Estado - 12 - Recibo del bien y/o prestación del servicio, por favor espere...";
                }
                else
                {
                    return idMessage + " - Downloading documents Status - 12 - Receipt of the good and/or provision of the service, please wait...";
                }
            }
            else if (idMessage == "FER10022")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - Descargando documentos Estado - 13 - Precargado, por favor espere...";
                }
                else
                {
                    return idMessage + " - Downloading documents Status - 13 - Preloaded, please wait...";
                }
            }
            else if (idMessage == "FER10023")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - No se encontraron documentos";
                }
                else
                {
                    return idMessage + " - No documents found";
                }
            }
            else if (idMessage == "FER10024")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + "- Documentos sincronizados correctamente.";
                }
                else
                {
                    return idMessage + " - Documents synchronized correctly.";
                }
            }
            else if (idMessage == "FER10025")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - Sincronizando documentos " + sComodinMessage[0] + " " + sComodinMessage[1] + " , Esto puede tomar bastante tiempo, por favor espere...";
                }
                else
                {
                    return idMessage + " - Synchronizing documents " + sComodinMessage[0] + " " + sComodinMessage[1] + ", This may take a long time, please wait... ";
                }
            }
            else if (idMessage == "FER10026")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - Por favor seleccionar la Fecha Inicial";
                }
                else
                {
                    return idMessage + "- Please select the Start Date";
                }
            }
            else if (idMessage == "FER10027")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - Por favor seleccionar la Fecha Final";
                }
                else
                {
                    return idMessage + "- Please select the End Date";
                }
            }
            else if (idMessage == "FER10028")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - Por favor parametrizar el nombre del usuario en el modulo recursos humanos.";
                }
                else
                {
                    return idMessage + "- Please parameterize the user name in the human resources module.";
                }
            }
            else if (idMessage == "FER10028")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - Por favor parametrizar el Apellido del usuario en el modulo recursos humanos.";
                }
                else
                {
                    return idMessage + "- Please parameterize the user's Last Name in the human resources module.";
                }
            }
            else if (idMessage == "FER10030")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - Por favor parametrizar el Cargo del usuario en el modulo recursos humanos.";
                }
                else
                {
                    return idMessage + "- Please parameterize the user's Position in the human resources module.";
                }
            }
            else if (idMessage == "FER10031")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - Por favor parametrizar el Departamento del usuario en el modulo recursos humanos.";
                }
                else
                {
                    return idMessage + " - Please parameterize the user's Department in the human resources module.";
                }
            }
            else if (idMessage == "FER10032")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - Por favor parametrizar el Numero de cedula del usuario en el modulo recursos humanos.";
                }
                else
                {
                    return idMessage + "- Please parameterize the user's ID number in the human resources module.";
                }
            }
            else if (idMessage == "FER10033")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {

                    return idMessage + " - Por favor parametrizar el Tipo de documento del usuario en el modulo recursos humanos.";
                }
                else
                {
                    return idMessage + " - Please parameterize the user's Document Type in the human resources module.";
                }
            }
            else if (idMessage == "FER10034")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {
                    return idMessage + " - Se cambio el estado del documento, esperar de 1 a 5 minutos para ver reflejado el cambio en la DIAN ";
                }
                else
                {
                    return idMessage + " - The status of the document was changed, wait 1 to 5 minutes to see the change reflected in the DIAN";
                }
            }
            else if (idMessage == "FER10035")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {
                    return idMessage + " - El usuario en SAP no esta autorizado para generar eventos en las facturas de proveedor ";
                }
                else
                {
                    return idMessage + " - The SAP is not authorized to generate events on supplier invoices.";
                }
            }
            else if (idMessage == "FER10036")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {
                    return idMessage + " - Consultando documentos, por favor esperar.";
                }
                else
                {
                    return idMessage + " - Consulting documents, please wait.";
                }
            }
            else if (idMessage == "FER10037")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {
                    return idMessage + " - Validando información, por favor espere...";
                }
                else
                {
                    return idMessage + " - Validating information, please wait...";
                }
            }
            else if (idMessage == "FER10038")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {
                    return idMessage + " - Por favor seleccione el evento que desea generar para poder continuar.";
                }
                else
                {
                    return idMessage + " - Please select the event you wish to generate to continue.";
                }
            }
            else if (idMessage == "FER10039")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {
                    return idMessage + " - Por favor seleccione almenos un documento para generar los eventos.";
                }
                else
                {
                    return idMessage + " - Please select at least one document to generate events.";
                }
            }
            else if (idMessage == "FER10040")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {
                    return idMessage + " - Por favor filtre por un estado DIAN especifico para poder generar eventos masivamente";
                }
                else
                {
                    return idMessage + " - Please filter by a specific DIAN state to be able to generate events in bulk.";
                }
            }
            else if (idMessage == "FER10041")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {
                    return idMessage + " - Actualizando eventos masivamente, por favor espere...";
                }
                else
                {
                    return idMessage + " - Updating events in bulk, please wait...";
                }
            }
            else if (idMessage == "FER10042")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {
                    return idMessage + " - Proceso finalizado correctamente.";
                }
                else
                {
                    return idMessage + " - Process completed successfully.";
                }
            }
            else if (idMessage == "FER10043")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {
                    return idMessage + " - Generando archivo, por favor espere...";
                }
                else
                {
                    return idMessage + " - Generating file, please wait...";
                }
            }
            else if (idMessage == "FER10044")
            {
                if (SboAppFacturacionElectronica.Language == BoLanguages.ln_Spanish_La)
                {
                    return idMessage + $" - Se genero el archivo {sComodinMessage[0]} en el escritorio correctamente. ";
                }
                else
                {
                    return idMessage + $" - The file {sComodinMessage[0]} was successfully generated on the desktop.";
                }
            }
            else
            {
                return "";
            }

            #endregion
        }

        public string VersionDll()
        {
            try
            {

                Assembly Assembly = Assembly.LoadFrom("BOElectronicReception.dll");
                Version vVersion = Assembly.GetName().Version;

                String VersionDll = vVersion.ToString();

                return VersionDll;
            }
            catch (Exception)
            {

                throw;
            }

        }

    }
}
