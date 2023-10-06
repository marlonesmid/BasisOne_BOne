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


namespace BOElectronicReception
{
    public class ElectronicReception
    {
        #region Instanciacion

        Funciones.Comunes DllFunciones = new Funciones.Comunes();

        #endregion

        #region Parametros globales TFHKA

        Recepcion21WS.ReceptorWSClient RecepcionWS;

        BasicHttpBinding port = new BasicHttpBinding();

        #endregion

        public void CreacionTablasyCamposeBillingBO(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {
                #region Creacion de tablas

                //1
                DllFunciones.ProgressBar(oCompany, sboapp, 27, 1, "Creando Tabla - Recepcion Electronica - Documentos recibidos PT, por favor espere...");
                DllFunciones.crearTabla(oCompany, sboapp, "BOTREDRPT", "BO Doc. Rec. Prov", SAPbobsCOM.BoUTBTableType.bott_NoObject);

                #endregion

                #region Creacion Campos

                DllFunciones.ProgressBar(oCompany, sboapp, 27, 1, "Creando Campo - Recep. Electro. - PT, por favor espere...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 10, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTPT", "Proveedor Technologico");

                DllFunciones.ProgressBar(oCompany, sboapp, 27, 1, "Creando Campo - Recep. Electro. - Numero Documento PT, por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 10, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTNDPT", "Numero Documento PT");

                DllFunciones.ProgressBar(oCompany, sboapp, 27, 1, "Creando Campo - Recep. Electro. - CUFE, por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTCUFE", "CUFE");

                DllFunciones.ProgressBar(oCompany, sboapp, 27, 1, "Creando Campo - Recep. Electro. - Status DIAN Codigo, por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 10, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTSDC", "Status DIAN Codi.");
                
                DllFunciones.ProgressBar(oCompany, sboapp, 27, 1, "Creando Campo - Recep. Electro. - Status DIAN Descripc., por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 10, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTSDD", "Status DIAN Desc.");

                DllFunciones.ProgressBar(oCompany, sboapp, 27, 1, "Creando Campo - Recep. Electro. - Status DIAN Fecha., por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 25, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTSDF", "Status DIAN Fech.");

                DllFunciones.ProgressBar(oCompany, sboapp, 27, 1, "Creando Campo - Recep. Electro. - Fecha Emision, por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 25, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTFE", "Fecha Emision.");

                DllFunciones.ProgressBar(oCompany, sboapp, 27, 1, "Creando Campo - Recep. Electro. - Fecha Recepcion, por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 25, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTFR", "Fecha Recep.");

                DllFunciones.ProgressBar(oCompany, sboapp, 27, 1, "Creando Campo - Recep. Electro. - Hora Emision, por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 25, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTHE", "Hora Emision.");

                DllFunciones.ProgressBar(oCompany, sboapp, 27, 1, "Creando Campo - Recep. Electro. - Monto Total, por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Float, BoFldSubTypes.st_Price, 11, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTMT", "Monto Total");

                DllFunciones.ProgressBar(oCompany, sboapp, 27, 1, "Creando Campo - Recep. Electro. - Numero Factura, por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 50, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTNF", "Numero Factura");

                DllFunciones.ProgressBar(oCompany, sboapp, 27, 1, "Creando Campo - Recep. Electro. - Numero Identificacion, por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 20, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTNI", "Numero Identificacion");

                DllFunciones.ProgressBar(oCompany, sboapp, 27, 1, "Creando Campo - Recep. Electro. - Razon Social, por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTNRS", "Razon Social");

                DllFunciones.ProgressBar(oCompany, sboapp, 27, 1, "Creando Campo - Recep. Electro. - Tipo Documento, por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 2, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTTD", "Tipo Documento");

                DllFunciones.ProgressBar(oCompany, sboapp, 27, 1, "Creando Campo - Recep. Electro. - Tipo Emisor, por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 2, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTTE", "Tipo Emisor");

                DllFunciones.ProgressBar(oCompany, sboapp, 27, 1, "Creando Campo - Recep. Electro. - Tipo Identidad, por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 3, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTTI", "Tipo Identidad");

                DllFunciones.ProgressBar(oCompany, sboapp, 27, 1, "Creando Campo - Estado Documento DIAN - Tipo Identidad, por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 3, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTEDD", "Est. Doc. DIAN");

                DllFunciones.ProgressBar(oCompany, sboapp, 27, 1, "Creando Campo - Numero Documento Preeliminar , por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 3, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTNDP", "Num. Doc. Preel");

                DllFunciones.ProgressBar(oCompany, sboapp, 27, 1, "Creando Campo - Numero Documento Definitivo , por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 3, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTNDD", "Num. Doc. Def.");

                DllFunciones.ProgressBar(oCompany, sboapp, 27, 1, "Creando Campo - Estado Evento DIAN , por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 3, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTEED", "Est. Even. DIAN");

                DllFunciones.ProgressBar(oCompany, sboapp, 27, 1, "Creando Campo - Path Adjunto XML , por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Memo, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTPAXML", "Path. Adju. XML");

                DllFunciones.ProgressBar(oCompany, sboapp, 27, 1, "Creando Campo - Codigo Respuesta WS XML , por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 10, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTCWSXML", "Cod. Resp. WS XML");

                DllFunciones.ProgressBar(oCompany, sboapp, 27, 1, "Creando Campo - Path Adjunto PDF , por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Memo, BoFldSubTypes.st_None, 254, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTPAPDF", "Path. Adju. PDF");

                DllFunciones.ProgressBar(oCompany, sboapp, 27, 1, "Creando Campo - Codigo Respuesta WS PDF , por favor espere ...");
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 10, "", BoYesNoEnum.tNO, null, "@BOTREDRPT", "BOTCWSPDF", "Cod. Resp. WS PDF");

                DllFunciones.ProgressBar(oCompany, sboapp, 27, 1, "Creando Campo - Tipo de documento , por favor espere ...");
                string[] ValidValuesFields1 = { "13", "Cedula Ciudadania", "31", "NIT" };
                DllFunciones.CreaCamposUsr(oCompany, sboapp, BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 10, "", BoYesNoEnum.tNO, ValidValuesFields1, "OHEM", "BOTTD", "Tipo documento");

                #endregion

                #region Crea procedimientos almacenados

                DllFunciones.ProgressBar(oCompany, sboapp, 27, 1, "Creando Procedimientos Almacenados Por favor espere ...");

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

        public void DescargaDocumentosTFHKA(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Form _oFormVisorRepcecion)
        {
            try
            {
                int iReprocesar = DllFunciones.sendMessageBoxY_N(_sboapp, "Se descargan los documentos de compras pendientes, ¿ Desea continuar ?");

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

                    EndpointAddress endPointEmision = new EndpointAddress(sURLEmision); //URL Recepcion

                    Recepcion21WS.ReceptorWSClient serviceClienTFHKAReception;

                    serviceClienTFHKAReception = new Recepcion21WS.ReceptorWSClient(port, endPointEmision);

                    ReceptorReporteStatusRequest ParametrosConsultaDocumentosRecepcion = new ReceptorReporteStatusRequest();

                    #endregion

                    #region Parametros Generales Reporte TFHKA

                    ParametrosConsultaDocumentosRecepcion.consecutivo = sConsecutivoTFHKA;
                    ParametrosConsultaDocumentosRecepcion.tokenEmpresa = sTokenEmpresa;
                    ParametrosConsultaDocumentosRecepcion.tokenPassword = sTokenPassword;

                    #endregion

                    #region Variables y Objetos

                    string sPathImages = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\BOElectronicReception\\Images\\";

                    #endregion

                    #region Consultas documentos por estado

                    DllFunciones.ProgressBar(_oCompany, _sboapp, 8, 1, "Descargando documentos Estado - 00 - Cargado, por favor espere...");
                    SincronizacionWS("00", "Cargado", ParametrosConsultaDocumentosRecepcion, serviceClienTFHKAReception, _oCompany, _sboapp, oSyncDocsRecep);


                    DllFunciones.ProgressBar(_oCompany, _sboapp, 8, 1, "Descargando documentos Estado - 01 - Entregado, por favor espere...");
                    SincronizacionWS("01", "Entregado", ParametrosConsultaDocumentosRecepcion, serviceClienTFHKAReception, _oCompany, _sboapp, oSyncDocsRecep);


                    DllFunciones.ProgressBar(_oCompany, _sboapp, 8, 1, "Descargando documentos Estado - 02 - Aceptación expresa (DIAN), por favor espere...");
                    SincronizacionWS("02", "Aceptación expresa (DIAN)", ParametrosConsultaDocumentosRecepcion, serviceClienTFHKAReception, _oCompany, _sboapp, oSyncDocsRecep);


                    DllFunciones.ProgressBar(_oCompany, _sboapp, 8, 1, "Descargando documentos Estado - 04 - Reclamo (DIAN), por favor espere...");
                    SincronizacionWS("04", "Reclamo (DIAN)", ParametrosConsultaDocumentosRecepcion, serviceClienTFHKAReception, _oCompany, _sboapp, oSyncDocsRecep);


                    DllFunciones.ProgressBar(_oCompany, _sboapp, 8, 1, "Descargando documentos Estado - 10 - Acuse de recibo (DIAN), por favor espere...");
                    SincronizacionWS("10", "Acuse de recibo (DIAN", ParametrosConsultaDocumentosRecepcion, serviceClienTFHKAReception, _oCompany, _sboapp, oSyncDocsRecep);


                    DllFunciones.ProgressBar(_oCompany, _sboapp, 8, 1, "Descargando documentos Estado - 11 - Rechazado (DIAN), por favor espere...");
                    SincronizacionWS("11", "Rechazado (DIAN)", ParametrosConsultaDocumentosRecepcion, serviceClienTFHKAReception, _oCompany, _sboapp, oSyncDocsRecep);


                    DllFunciones.ProgressBar(_oCompany, _sboapp, 8, 1, "Descargando documentos Estado - 12 - Recibo del bien y/o prestación del servicio, por favor espere...");
                    SincronizacionWS("12", "Recibo del bien y/o prestación del servicio", ParametrosConsultaDocumentosRecepcion, serviceClienTFHKAReception, _oCompany, _sboapp, oSyncDocsRecep);


                    DllFunciones.ProgressBar(_oCompany, _sboapp, 8, 1, "Descargando documentos Estado - 13 - Precargado, por favor espere...");
                    SincronizacionWS("13", "Precargado", ParametrosConsultaDocumentosRecepcion, serviceClienTFHKAReception, _oCompany, _sboapp, oSyncDocsRecep);

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
                    SAPbouiCOM.ComboBox oEstado = (SAPbouiCOM.ComboBox)_oFormVisorRepcecion.Items.Item("cboStado").Specific;

                    SAPbobsCOM.Recordset oRecorsetInvoices = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    SAPbobsCOM.Recordset oRecorsetCreditMemo = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    SAPbobsCOM.Recordset oRecorsetDebitMemo = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    SAPbouiCOM.DataTable oTableInvoices = _oFormVisorRepcecion.DataSources.DataTables.Item("DT_Invoices");
                    SAPbouiCOM.DataTable oTableCreditMemo = _oFormVisorRepcecion.DataSources.DataTables.Item("DT_CreditMemo");
                    SAPbouiCOM.DataTable oTableDebitMemo = _oFormVisorRepcecion.DataSources.DataTables.Item("DT_DebitMemo");

                    sInvoices = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetInvoices");
                    sInvoices = sInvoices.Replace("%PathImages%", sPathImages).Replace("%FI%","20220108").Replace("%FF%","20251231");

                    sCreditMemo = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetCreditMemo");
                    sInvoices = sInvoices.Replace("%PathImages%", sPathImages);

                    sDebitMemo = DllFunciones.GetStringXMLDocument(_oCompany, "BOElectronicReception", "ElectronicReception", "GetDebitMemo");
                    sInvoices = sInvoices.Replace("%PathImages%", sPathImages);

                    if (oEstado.Value == "-" || oEstado.Value == "")
                    {
                        sInvoices = sInvoices.Replace("%Estado%", "");
                    }
                    else
                    {
                        sInvoices = sInvoices.Replace("%Estado%", "AND \"U_BOTSDC\" = '" + oEstado.Value + "' ");
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
                }

                

            }
            catch (Exception)
            {

                throw;
            }
        }

        public void LoadFormDocumentsReception(SAPbouiCOM.Application sboapp, SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oFormReception)
        {
            Funciones.Comunes DLLFunciones = new Funciones.Comunes();

            try
            {
                #region Variables y objetos

                //SAPbouiCOM.PictureBox oLogoBO = (SAPbouiCOM.PictureBox)oFormReception.Items.Item("imgLogoBO").Specific;

                SAPbouiCOM.Folder oFolder1 = (SAPbouiCOM.Folder)oFormReception.Items.Item("Folder1").Specific;

                #endregion

                #region Centra en pantalla formulario

                oFormReception.Left = (sboapp.Desktop.Width - oFormReception.Width) / 2;
                oFormReception.Top = (sboapp.Desktop.Height - oFormReception.Height) / 4;

                #endregion

                #region Asignacion Logo BO

                //oLogoBO.Picture = (Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Core\\Imagenes\\BO.jpg");

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

        public void SincronizacionWS(string _sCodigoStatusDIAN, string _DescripcionEstatusDIAN, ReceptorReporteStatusRequest ParametrosConsultaDocumentosRecepcion, Recepcion21WS.ReceptorWSClient _serviceClienTFHKAReception, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Application sboapp, SAPbobsCOM.Recordset oSyncDocsRecep)
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

                    

                    oSyncDocsRecep.DoQuery(sSyncDocsRecepCopia);

                }

                #endregion
            }

        }

        public void DescargaXML_PDF(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Form _oFormVisorRepcecion)
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

        public void MatrixOpenFile(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Form _oFormVDR, ItemEvent pVal, string _TipoDocumento, string _ColUID)
        {
            if (_TipoDocumento == "XML" && _ColUID == "Col_20")
            {
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
            }

            if (_TipoDocumento == "PDF" && _ColUID == "Col_21")
            {
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

                
            }
        }

        public void ChagueStatusDocument(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Form _oFormVisorRepcecion, ItemEvent pVal, string _TipoEventoDIAN, string _ColUID)
        {
            if (_ColUID == "Col_22")
            {
                #region Aceptacion documento

              
                SAPbouiCOM.Matrix oMatrixOPCH = (Matrix)_oFormVisorRepcecion.Items.Item("MtxOPCH").Specific;

                if (pVal.Row == 0)
                {

                }
                else
                {
                    #region Variables y Objetos

                    string sNumeroDocumentoFacturaProveedor = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific)).Value;
                    string sIdentificacionEmisor = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific)).Value;
                    string sNombreProveedor = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_3").Cells.Item(pVal.Row).Specific)).Value;

                    #endregion

                    int iProcesar = DllFunciones.sendMessageBoxY_N(_sboapp, "Se aceptara el documento " + sNumeroDocumentoFacturaProveedor + " del proveedor " + sNombreProveedor +" . ¿ Desea Continuar ? ");

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
                                DllFunciones.sendMessageBox(_sboapp, "RE001 - Por favor parametrizar el nombre del usuario en el modulo recursos humanos.");
                            }
                            else if (string.IsNullOrEmpty(sApellidoAceptador))
                            {
                                DllFunciones.sendMessageBox(_sboapp, "RE002 - Por favor parametrizar el Apellido del usuario en el modulo recursos humanos.");
                            }
                            else if (string.IsNullOrEmpty(sCargoAceptador))
                            {
                                DllFunciones.sendMessageBox(_sboapp, "RE003 - Por favor parametrizar el Cargo del usuario en el modulo recursos humanos.");
                            }
                            else if (string.IsNullOrEmpty(sDepartamentoAceptador))
                            {
                                DllFunciones.sendMessageBox(_sboapp, "RE004 - Por favor parametrizar el Departamento del usuario en el modulo recursos humanos.");
                            }
                            else if (string.IsNullOrEmpty(sNITAceptador))
                            {
                                DllFunciones.sendMessageBox(_sboapp, "RE005 - Por favor parametrizar el Numero de cedula del usuario en el modulo recursos humanos.");
                            }
                            else if (string.IsNullOrEmpty(sTipoDocumentoAceptador))
                            {
                                DllFunciones.sendMessageBox(_sboapp, "RE006 - Por favor parametrizar el Tipo de documento del usuario en el modulo recursos humanos.");
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
                                ParametrosCambioEstatus.tipoIdentificacionemisor = sTipoDocumentoAceptador;
                                ParametrosCambioEstatus.numeroDocumento = sNumeroDocumentoFacturaProveedor;

                                ReceptorCambioEstatusRequest.EjecutadoPorRequest UsuarioAceptador = new ReceptorCambioEstatusRequest.EjecutadoPorRequest();

                                UsuarioAceptador.Nombre = sNombreAceptador;
                                UsuarioAceptador.Apellido = sApellidoAceptador;
                                UsuarioAceptador.Cargo = sCargoAceptador;
                                UsuarioAceptador.Departamento = sDepartamentoAceptador;
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

                                ParametrosCambioEstatus.status = "10";
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

                                    sUpdateStatusDocument = sUpdateStatusDocument.Replace("%NumeroFactura%", sNumeroDocumentoFacturaProveedor).Replace("%NumeroIdentificacion%", sIdentificacionEmisor);

                                    oUpdateStatusDocument.DoQuery(sUpdateStatusDocument);

                                    #endregion
                                }
                                else
                                {
                                    DllFunciones.sendMessageBox(_sboapp, "RE007 - No se pudo cambiar el estaodo" + wsCambiarEstado.mensaje.ToString());
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
                                    sInvoices = sInvoices.Replace("%Estado%", "AND \"U_BOTSDC\" = '" + oEstado.Value + "' ");
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

                            }

                            #endregion

                        }
                        else
                        {


                        }
                    }



                }

                #endregion
            }
            else if (_ColUID == "Col_7")
            {
                #region Aceptacion documento

                SAPbouiCOM.Matrix oMatrixOPCH = (Matrix)_oFormVisorRepcecion.Items.Item("MtxOPCH").Specific;

                if (pVal.Row == 0)
                {

                }
                else
                {
                    #region Variables y Objetos

                    string sNumeroDocumentoFacturaProveedor = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific)).Value;
                    string sIdentificacionEmisor = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific)).Value;
                    string sNombreProveedor = ((SAPbouiCOM.EditText)(oMatrixOPCH.Columns.Item("Col_3").Cells.Item(pVal.Row).Specific)).Value;

                    #endregion

                    int iProcesar = DllFunciones.sendMessageBoxY_N(_sboapp, "Se rechazara el documento " + sNumeroDocumentoFacturaProveedor + " del proveedor " + sNombreProveedor + " . ¿ Desea Continuar ? ");

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
                                DllFunciones.sendMessageBox(_sboapp, "RE001 - Por favor parametrizar el nombre del usuario en el modulo recursos humanos.");
                            }
                            else if (string.IsNullOrEmpty(sApellidoAceptador))
                            {
                                DllFunciones.sendMessageBox(_sboapp, "RE002 - Por favor parametrizar el Apellido del usuario en el modulo recursos humanos.");
                            }
                            else if (string.IsNullOrEmpty(sCargoAceptador))
                            {
                                DllFunciones.sendMessageBox(_sboapp, "RE003 - Por favor parametrizar el Cargo del usuario en el modulo recursos humanos.");
                            }
                            else if (string.IsNullOrEmpty(sDepartamentoAceptador))
                            {
                                DllFunciones.sendMessageBox(_sboapp, "RE004 - Por favor parametrizar el Departamento del usuario en el modulo recursos humanos.");
                            }
                            else if (string.IsNullOrEmpty(sNITAceptador))
                            {
                                DllFunciones.sendMessageBox(_sboapp, "RE005 - Por favor parametrizar el Numero de cedula del usuario en el modulo recursos humanos.");
                            }
                            else if (string.IsNullOrEmpty(sTipoDocumentoAceptador))
                            {
                                DllFunciones.sendMessageBox(_sboapp, "RE006 - Por favor parametrizar el Tipo de documento del usuario en el modulo recursos humanos.");
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
                                ParametrosCambioEstatus.tipoIdentificacionemisor = sTipoDocumentoAceptador;
                                ParametrosCambioEstatus.numeroDocumento = sNumeroDocumentoFacturaProveedor;

                                ReceptorCambioEstatusRequest.EjecutadoPorRequest UsuarioAceptador = new ReceptorCambioEstatusRequest.EjecutadoPorRequest();

                                UsuarioAceptador.Nombre = sNombreAceptador;
                                UsuarioAceptador.Apellido = sApellidoAceptador;
                                UsuarioAceptador.Cargo = sCargoAceptador;
                                UsuarioAceptador.Departamento = sDepartamentoAceptador;
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

                                ParametrosCambioEstatus.status = "01";
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

                                    sUpdateStatusDocument = sUpdateStatusDocument.Replace("%NumeroFactura%", sNumeroDocumentoFacturaProveedor).Replace("%NumeroIdentificacion%", sIdentificacionEmisor);

                                    oUpdateStatusDocument.DoQuery(sUpdateStatusDocument);

                                    #endregion
                                }
                                else
                                {
                                    DllFunciones.sendMessageBox(_sboapp, "RE007 - No se pudo cambiar el estado - " + wsCambiarEstado.mensaje.ToString());
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
                                    sInvoices = sInvoices.Replace("%Estado%", "AND \"U_BOTSDC\" = '" + oEstado.Value + "' ");
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

                            }

                            #endregion

                        }
                        else
                        {


                        }
                    }



                }

                #endregion
            }
            else
            {
                DllFunciones.sendMessageBox(_sboapp, "Por favor parametrizar el empleado de ventas en el modulo de usuarios");
            }
        }

        public void LoadMatrixReception(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Form _oFormVisorRepcecion)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            SAPbouiCOM.EditText oFI = (SAPbouiCOM.EditText)_oFormVisorRepcecion.Items.Item("txtFI").Specific;
            SAPbouiCOM.EditText oFF = (SAPbouiCOM.EditText)_oFormVisorRepcecion.Items.Item("txtFF").Specific;
            SAPbouiCOM.ComboBox oEstado = (SAPbouiCOM.ComboBox)_oFormVisorRepcecion.Items.Item("cboStado").Specific;

            if (string.IsNullOrEmpty(oFI.Value))
            {
                DllFunciones.sendMessageBox(_sboapp, "Por favor seleccionar la Fecha Inicial");
            }
            else if (string.IsNullOrEmpty(oFF.Value))
            {
                DllFunciones.sendMessageBox(_sboapp, "Por favor seleccionar la Fecha Final");
            }
            else
            {
                #region Carga Infortmacion en la Matrix

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

                #region Liberar Objetos

                DllFunciones.liberarObjetos(oRecorsetCreditMemo);
                DllFunciones.liberarObjetos(oRecorsetDebitMemo);

                #endregion
            }

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
