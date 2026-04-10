using System;
using SAPbobsCOM;
using SAPbouiCOM;
using System.IO;
using System.Reflection;
using Presupuesto;
using System.Xml;
using Intercompany;
using System.Security.Cryptography;
using System.Text;
using BOProduccion;
using BOElectronicReception;
using BOCore;
using BOBillingConsolidator;

namespace BasisOne
{
    internal class SAPMain
    {
        #region Variables y objetos globales

        static Application sboapp;
        static SAPbobsCOM.Company _company;

        static bool Flag1 = false;
        public bool Flag2 = false;
        static bool FlagNew = false;

        public bool TieneLicenciaIntercompany = false;
        public bool TieneLicenciaPresupuesto = false;
        public bool TieneLicenciaeBilling = false;
        public bool TieneLicenciaProduction = false;
        public bool TieneLicenciaElectronicRepception = false;
        public bool TieneLicenciaConsolidadorFacturacion = false;
        public int IdLanguage = 0;
        
        string Llave = "B4s1s0neS4S";
        string sMotor = null;
        string sNameDB = null;
        string sPath = null;
        string sCurrentUser;
        string sPrefijoDocNumSM = null;
        

        #endregion

        #region Instanciacion Dll's

        Funciones.Comunes DllFunciones = new Funciones.Comunes();
        eBilling.eBillingBO DlleBilling = new eBilling.eBillingBO(sboapp, _company);
        Intercompany.Intercompany DllIntercompany = new Intercompany.Intercompany(sboapp, _company);
        BOProduccion.Production DllProduction = new BOProduccion.Production();
        BOElectronicReception.ElectronicReception DllElectronicReception = new ElectronicReception();
        BOCore.Core DllCore = new BOCore.Core();

        #endregion

        internal void Init()
        {            

            SAPbouiCOM.SboGuiApi guiApi = new SAPbouiCOM.SboGuiApi();
            String connectionString = Environment.GetCommandLineArgs().Length != 1 ? Environment.GetCommandLineArgs().GetValue(1).ToString() : "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";
            guiApi.Connect(connectionString);
            sboapp = guiApi.GetApplication();

            _company = (SAPbobsCOM.Company)sboapp.Company.GetDICompany();

            sNameDB = _company.CompanyDB;
            sMotor = Convert.ToString(_company.DbServerType);

            DllCore.TablasyCamposBaseBO(sboapp, _company, sNameDB);

            DllFunciones.StatusBar(sboapp, BoStatusBarMessageType.smt_Warning, DllCore.MessageSystem("00001", sboapp));

            //Crear menus        
            setMenus();

            //Handles por cierre de la aplicacion
            setEvents();

            //asignar los filtros en los eventos que debe estar a la escucha
            setFilter();

            
            DllFunciones.StatusBar(sboapp, BoStatusBarMessageType.smt_Success, DllCore.MessageSystem("00002", sboapp));

        }

        private void setMenus()
        {
            try
            {
                //  borrar el menú
                if (sboapp.Menus.Exists("mnuBasisOne"))
                {
                    sboapp.Menus.RemoveEx("mnuBasisOne");
                }

                #region Menu principal Basis One

                SAPbouiCOM.Menus oMenus = null;
                SAPbouiCOM.MenuItem oMenuItem = null;
                oMenus = sboapp.Menus;
                SAPbouiCOM.MenuCreationParams oCreationPackage = null;
                oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(sboapp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
                oMenuItem = sboapp.Menus.Item("43520");
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                oCreationPackage.UniqueID = "mnuBasisOne";
                oCreationPackage.String = "BOne";
                oCreationPackage.Enabled = true;
                oCreationPackage.Position = -1;
                oMenus = oMenuItem.SubMenus;

                #endregion

                try
                {
                    oMenus.AddEx(oCreationPackage);
                }
                catch (Exception e)
                {
                    DllFunciones.sendErrorMessage(sboapp, e);
                }

                try
                {
                    oMenuItem = sboapp.Menus.Item("mnuBasisOne");
                    oMenus = oMenuItem.SubMenus;

                    #region Menu Modulo Gestor-AddOn

                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    oCreationPackage.UniqueID = "mnuBO_GestorAddOn";
                    oCreationPackage.String = "Gestor AddOn";
                    oCreationPackage.Image = "";
                    oCreationPackage.Position = oCreationPackage.Position + 1;
                    oMenus.AddEx(oCreationPackage);

                    #endregion

                    AdicionSubmenu(oCreationPackage, oMenus, oMenuItem, sboapp, "AddInPresupuesto", sNameDB);

                    AdicionSubmenu(oCreationPackage, oMenus, oMenuItem, sboapp, "AddInIntercompany", sNameDB);

                    AdicionSubmenu(oCreationPackage, oMenus, oMenuItem, sboapp, "AddIneBillingBO", sNameDB);

                    AdicionSubmenu(oCreationPackage, oMenus, oMenuItem, sboapp, "AddInProduccion", sNameDB);

                    AdicionSubmenu(oCreationPackage, oMenus, oMenuItem, sboapp, "AddInElectronicReception", sNameDB);

                    AdicionSubmenu(oCreationPackage, oMenus, oMenuItem, sboapp, "AddInBillingConsolidator", sNameDB);

                }
                catch (Exception e)
                {
                    DllFunciones.sendErrorMessage(sboapp, e);
                }

                if (sboapp.Menus.Exists("mnuBasisOne"))
                {
                    sboapp.Menus.Item("mnuBasisOne").Image = (Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Core\\Images\\LogoBO20x20.bmp");
                }
            }
            catch (Exception e)
            {
                DllFunciones.sendErrorMessage(sboapp, e);
            }
        }

        private void setEvents()
        {
            ApplicationHandler ah = new ApplicationHandler(sboapp);

            sboapp.ItemEvent += Sboapp_ItemEvent;
            sboapp.ItemEvent += DlleBilling.ItemEvent_eBilling;
            sboapp.MenuEvent += Sboapp_MenuEvent;
            sboapp.AppEvent += ah.app_Handler;
            sboapp.RightClickEvent += Sboapp_RightClick;
            sboapp.FormDataEvent += Sboapp_DataEvent;

        }

        private void setFilter()
        {

            SAPbouiCOM.EventFilters oFilters = sboapp.GetFilter();
            if (oFilters == null)
            {
                oFilters = new SAPbouiCOM.EventFilters();
                sboapp.SetFilter(oFilters);
            }
            SAPbouiCOM.EventFilter oFilter1 = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK);
            oFilter1.AddEx("BO_ConfPresup");
            SAPbouiCOM.EventFilter oFilter2 = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD);
            oFilter2.AddEx("BOPC");

        }

        private void Sboapp_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                Presupuesto.Core oCore = new Presupuesto.Core(sboapp, _company);
                Funciones.Comunes oFunc = new Funciones.Comunes();
                Presupuesto.Core oPresup = new Presupuesto.Core(sboapp, _company);
                Intercompany.Intercompany DllIntercompany = new Intercompany.Intercompany(sboapp, _company);

                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_CLICK:

                        #region Eventos Gestor de AddOn

                        if (pVal.FormUID == "BO_Gestion_AddOn" && pVal.ItemUID == "lblUpdFea" && pVal.Before_Action == false && pVal.Action_Success == true )
                        {
                            #region Abre archivo actualizaciones y mejoras

                            SAPbouiCOM.Form oInvoice;

                            oInvoice = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                            DllCore.OpenFileUpdatesAndImprovements(sboapp, _company);                            

                            #endregion
                        }
                        else if (pVal.FormUID == "BO_Gestion_AddOn" && pVal.ItemUID == "lblHelp" && pVal.Before_Action == false && pVal.Action_Success == true)
                        {
                            #region Abre archivo actualizaciones y mejoras

                            SAPbouiCOM.Form oInvoice;

                            oInvoice = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                            DllCore.OpenURLHelpDesk(sboapp, _company);

                            #endregion
                        }
                        #endregion

                        if (TieneLicenciaPresupuesto == true)
                        {
                            #region Eventos_Presupuesto

                            if (pVal.FormUID == "BO_ConfPresup" && pVal.ItemUID == "Btn_Crea" && pVal.Before_Action == true && pVal.Action_Success == false)
                            {
                                oPresup.creaTablasPresup(_company, sboapp);
                            }
                            if (pVal.FormUID == "BOPC" && pVal.ItemUID == "1" && pVal.Before_Action == true && pVal.Action_Success == false)
                            {
                                BubbleEvent = oPresup.presupEventos(pVal.EventType, pVal);
                            }

                            #endregion
                        }

                        if (TieneLicenciaeBilling == true)
                        {
                            #region Eventos eBillingBO

                            #region Parametros eBilling

                            if (pVal.FormUID == "BO_eBillingP" && pVal.ItemUID == "1" && pVal.Before_Action == true && pVal.Action_Success == false)
                            {
                                #region Valida informacion del formulario parametros iniciales eBilling 

                                if (pVal.FormMode == 3 || pVal.FormMode == 2)
                                {

                                    SAPbouiCOM.Form oFormBO_eBillingP = sboapp.Forms.Item("BO_eBillingP");

                                    BubbleEvent = DlleBilling.UpdateDataFormParametrizacionEbilling(sboapp, _company, oFormBO_eBillingP, sMotor);

                                }

                                #endregion
                            }
                            else if (pVal.FormUID == "BO_eBillingP" && pVal.ItemUID == "chkCCEM" && pVal.Before_Action == true && pVal.Action_Success == false)
                            {
                                #region Valida informacion del formulario parametros iniciales eBilling 

                                DlleBilling.EnableButtonCreacionCamposEntregaMercancia(sboapp, _company);

                                #endregion
                            }
                            else if (pVal.FormUID == "BO_eBillingP" && pVal.ItemUID == "chkCCC" && pVal.Before_Action == true && pVal.Action_Success == false)
                            {
                                #region Valida informacion del formulario parametros iniciales eBilling 

                                DlleBilling.CheckBoxCloudControlCenterAndExportInvoice(sboapp, _company);

                                #endregion
                            }
                            #endregion

                            #region Visor documentos Facturacion Electronica

                            else if (pVal.FormUID == "BOVDEB" && pVal.ItemUID == "MtxOINV" && pVal.Before_Action == true && pVal.Action_Success == false)
                            {
                                #region Selecciona la infomracion de la matrix facturas visor de documentos

                                try
                                {
                                    #region Variables

                                    SAPbouiCOM.Form oForm_BOVDEB = sboapp.Forms.Item("BOVDEB");
                                    SAPbouiCOM.Matrix oMtxOINV = (Matrix)oForm_BOVDEB.Items.Item("MtxOINV").Specific;

                                    #endregion

                                    DllFunciones.SelectRowMatrix(oMtxOINV, pVal);

                                }
                                catch (Exception)
                                {

                                    throw;
                                }

                                #endregion
                            }
                            else if (pVal.FormUID == "BOVDEB" && pVal.ItemUID == "MtxORIN" && pVal.Before_Action == true && pVal.Action_Success == false)
                            {
                                #region Selecciona la infomracion de la matrix notas credito  visor de documentos

                                try
                                {
                                    #region Variables

                                    SAPbouiCOM.Form oForm_BOVDEB = sboapp.Forms.Item("BOVDEB");
                                    SAPbouiCOM.Matrix MtxORIN = (Matrix)oForm_BOVDEB.Items.Item("MtxORIN").Specific;

                                    #endregion

                                    DllFunciones.SelectRowMatrix(MtxORIN, pVal);

                                }
                                catch (Exception)
                                {

                                    throw;
                                }

                                #endregion
                            }
                            else if (pVal.FormUID == "BOVDEB" && pVal.ItemUID == "MtxOINVD" && pVal.Before_Action == true && pVal.Action_Success == false)
                            {
                                #region Selecciona la infomracion de la matrix notas debito visor de documentos

                                try
                                {
                                    #region Variables

                                    SAPbouiCOM.Form oForm_BOVDEB = sboapp.Forms.Item("BOVDEB");
                                    SAPbouiCOM.Matrix MtxOINVD = (Matrix)oForm_BOVDEB.Items.Item("MtxOINVD").Specific;

                                    #endregion

                                    DllFunciones.SelectRowMatrix(MtxOINVD, pVal);

                                }
                                catch (Exception)
                                {

                                    throw;
                                }

                                #endregion
                            }
                            else if (pVal.FormUID == "BOVDEB" && pVal.ItemUID == "MtxOPCH" && pVal.Before_Action == true && pVal.Action_Success == false)
                            {
                                #region Selecciona la infomracion de la matrix notas debito visor de documentos

                                try
                                {
                                    #region Variables

                                    SAPbouiCOM.Form oForm_BOVDEB = sboapp.Forms.Item("BOVDEB");
                                    SAPbouiCOM.Matrix MtxOINVD = (Matrix)oForm_BOVDEB.Items.Item("MtxOPCH").Specific;

                                    #endregion

                                    DllFunciones.SelectRowMatrix(MtxOINVD, pVal);

                                }
                                catch (Exception)
                                {

                                    throw;
                                }

                                #endregion
                            }

                            #endregion

                            #region Socios de negocio

                            else if (pVal.FormType == 134 && pVal.ItemUID == "1" && pVal.Before_Action == true && pVal.Action_Success == false && (pVal.FormMode == 2 || pVal.FormMode == 3))
                            {
                                #region Valida la informacion de los socios de negocio 

                                SAPbouiCOM.Form oBusinessPartnerd;

                                oBusinessPartnerd = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                BubbleEvent = DlleBilling.Validate_oBusinessPartnerd(sboapp, _company, oBusinessPartnerd, sMotor);

                                #endregion

                            }

                            #endregion

                            #region Factura de venta

                            else if (pVal.FormType == 133 && pVal.ItemUID == "1" && pVal.Before_Action == true && pVal.Action_Success == false && (pVal.FormMode == 2 || pVal.FormMode == 3))
                            {
                                #region Valida la informacion de las facturas de venta

                                SAPbouiCOM.Form oInvoice;

                                oInvoice = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                BubbleEvent = DlleBilling.Validate_oInvoices(sboapp, _company, oInvoice, sMotor);

                                #endregion

                            }
                            else if (pVal.FormType == 133 && pVal.ItemUID == "lblURL" && pVal.Before_Action == false && pVal.Action_Success == true && (pVal.FormMode == 1 || pVal.FormMode == 2))
                            {
                                #region abre la factura en la DIAN

                                SAPbouiCOM.Form oInvoice;

                                oInvoice = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                DlleBilling.DocumentSearchDIAN(oInvoice);

                                #endregion

                            }
                            else if (pVal.FormType == 141 && pVal.ItemUID == "lblURL" && pVal.Before_Action == false && pVal.Action_Success == true && (pVal.FormMode == 1 || pVal.FormMode == 2))
                            {
                                #region abre la factura en la DIAN

                                SAPbouiCOM.Form oInvoice;

                                oInvoice = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                DlleBilling.DocumentSearchDIAN(oInvoice);

                                #endregion

                            }

                            #endregion

                            #region Factura + Pago

                            else if (pVal.FormType == 60090 && pVal.ItemUID == "1" && pVal.Before_Action == true && pVal.Action_Success == false && (pVal.FormMode == 2 || pVal.FormMode == 3))
                            {
                                #region Valida la informacion de las facturas de anticipo

                                SAPbouiCOM.Form oInvoice_Payment;

                                oInvoice_Payment = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                BubbleEvent = DlleBilling.Validate_oInvoices(sboapp, _company, oInvoice_Payment, sMotor);

                                #endregion

                            }
                            else if (pVal.FormType == 60090 && pVal.ItemUID == "lblURL" && pVal.Before_Action == false && pVal.Action_Success == true && (pVal.FormMode == 1 || pVal.FormMode == 2))
                            {
                                #region Abre la factura en la DIAN

                                SAPbouiCOM.Form oInvoicePayment;

                                oInvoicePayment = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                DlleBilling.DocumentSearchDIAN(oInvoicePayment);

                                #endregion

                            }

                            #endregion

                            #region Factura de proveedores - Documento Soporte

                            else if (pVal.FormType == 141 && pVal.ItemUID == "1" && pVal.Before_Action == true && pVal.Action_Success == false && (pVal.FormMode == 2 || pVal.FormMode == 3))
                            {
                                #region Valida la informacion de las facturas de anticipo

                                SAPbouiCOM.Form oInvoice_Payment;

                                oInvoice_Payment = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);                                

                                #endregion

                            }


                            #endregion

                            #region Nota credito de cliente

                            else if (pVal.FormType == 179 && pVal.ItemUID == "1" && pVal.Before_Action == true && pVal.Action_Success == false && (pVal.FormMode == 2 || pVal.FormMode == 3))
                            {
                                #region Valida la informacion de las notas credito de cliente 

                                SAPbouiCOM.Form oCreditNote;

                                oCreditNote = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                BubbleEvent = DlleBilling.Validate_oCreditNote(sboapp, _company, oCreditNote, sMotor);

                                #endregion

                            }
                            else if (pVal.FormType == 179 && pVal.ItemUID == "lblURL" && pVal.Before_Action == false && pVal.Action_Success == true && (pVal.FormMode == 1 || pVal.FormMode == 2))
                            {
                                #region Abre la factura en la DIAN

                                SAPbouiCOM.Form oCreditNote;

                                oCreditNote = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                DlleBilling.DocumentSearchDIAN(oCreditNote);

                                #endregion
                            }

                            #endregion

                            #region Notas debido de cliente

                            else if (pVal.FormType == 65303 && pVal.ItemUID == "1" && pVal.Before_Action == true && pVal.Action_Success == false && (pVal.FormMode == 2 || pVal.FormMode == 3))
                            {
                                #region Valida la informacion de las notas debito de cliente

                                SAPbouiCOM.Form oDebitNote;

                                oDebitNote = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                BubbleEvent = DlleBilling.Validate_oDebitNote(sboapp, _company, oDebitNote, sMotor);

                                #endregion
                            }

                            #endregion

                            #region Nueva orden de produccion 

                            else if (pVal.FormUID == "BO_New_WO" && pVal.ItemUID == "mtxRP" && pVal.Before_Action == true && pVal.Action_Success == false)
                            {
                                #region Selecciona la infomracion de la matrix Nueva orden de produccion

                                try
                                {
                                    #region Variables

                                    SAPbouiCOM.Form BO_New_WO = sboapp.Forms.Item("BO_New_WO");
                                    SAPbouiCOM.Matrix oMtxNWO = (Matrix)BO_New_WO.Items.Item("mtxRP").Specific;

                                    #endregion

                                    DllFunciones.SelectRowMatrix(oMtxNWO, pVal);

                                }
                                catch (Exception)
                                {

                                    throw;
                                }

                                #endregion
                            }

                            #endregion

                            #endregion
                        }

                        if (TieneLicenciaProduction == true)
                        {
                            #region Eventos Produccion

                            if (pVal.FormUID == "BOFormCOP" && pVal.ItemUID == "MtxCOP" && pVal.Before_Action == true && pVal.Action_Success == false)
                            {
                                try
                                {
                                    #region Variables

                                    SAPbouiCOM.Form oForm_COP = sboapp.Forms.Item("BOFormCOP");
                                    SAPbouiCOM.Matrix MtxCOP = (Matrix)oForm_COP.Items.Item("MtxCOP").Specific;

                                    #endregion

                                    DllFunciones.SelectRowMatrix(MtxCOP, pVal);

                                }
                                catch (Exception)
                                {

                                    throw;
                                }
                            }

                            else if (pVal.FormUID == "BOFormCOP" && pVal.ColUID == "Col_19" && pVal.ItemUID == "MtxCOP" && pVal.BeforeAction == false)
                            {
                                SAPbouiCOM.Form oFormCOP = sboapp.Forms.Item("BOFormCOP");

                                DllProduction.LoadFormMRawMaterial(sboapp, _company, oFormCOP, pVal);

                            }
                            if (pVal.FormUID == "BOFormMPC" && pVal.ItemUID == "MtxMPE" && pVal.Before_Action == true && pVal.Action_Success == false)
                            {
                                try
                                {
                                    #region Variables

                                    SAPbouiCOM.Form oForm_COP = sboapp.Forms.Item("BOFormMPC");
                                    SAPbouiCOM.Matrix MtxCOP = (Matrix)oForm_COP.Items.Item("MtxMPE").Specific;

                                    #endregion

                                    DllFunciones.SelectRowMatrix(MtxCOP, pVal);

                                }
                                catch (Exception)
                                {

                                    throw;
                                }
                            }

                            #endregion
                        }

                        if (TieneLicenciaElectronicRepception == true)
                        {
                            #region Eventos Recepcion Electronica 

                            if (pVal.FormUID == "BOTVDR" && pVal.ItemUID == "MtxOPCH" && pVal.Before_Action == true && pVal.Action_Success == false)
                            {
                                #region Selecciona la infomracion de la matrix facturas visor de documentos

                                try
                                {
                                    #region Variables

                                    SAPbouiCOM.Form oForm_BOVDEB = sboapp.Forms.Item("BOTVDR");
                                    SAPbouiCOM.Matrix oMtxOINV = (Matrix)oForm_BOVDEB.Items.Item("MtxOPCH").Specific;

                                    #endregion

                                    DllFunciones.SelectRowMatrix(oMtxOINV, pVal);

                                }
                                catch (Exception)
                                {

                                    throw;
                                }

                                #endregion
                            }
                            #endregion
                        }

                        break;

                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:

                        if (TieneLicenciaPresupuesto == true)
                        {
                            #region Eventos_Presupuesto
                            if (pVal.FormUID == "BOPC" && pVal.ItemUID == "BO_User" && pVal.BeforeAction == false)
                            {
                                //marcar salida del formulario de seleccion de usuario para poder llenar campo con id de usuario que se usas en el drilldown
                                Flag1 = true;
                            }
                            #endregion
                        }

                        if (TieneLicenciaeBilling == true)
                        {
                            #region Eventos_eBilling

                            if (pVal.FormUID == "BOVDEB" && pVal.ItemUID == "txtSN" && pVal.BeforeAction == false)
                            {
                                SAPbouiCOM.Form oFormVisorDocs = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                DlleBilling.ChooFormListSN(FormUID, oFormVisorDocs, pVal, _company, sboapp);
                            }

                            #endregion
                        }

                        if (TieneLicenciaProduction == true)
                        {
                            #region Eventos_Produccion

                            if (pVal.FormUID == "BO_New_WO" && pVal.ItemUID == "mtxRP" && pVal.BeforeAction == false)
                            {
                                SAPbouiCOM.Form oForWO = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                DllProduction.CFLAfterMatrix(FormUID, oForWO, pVal, _company, sboapp);

                            }
                            else if (pVal.FormUID == "BO_New_WO" && pVal.ItemUID == "txtIPR" && pVal.BeforeAction == false)
                            {
                                SAPbouiCOM.Form oForWO = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                DllProduction.CFLAfterIRP(FormUID, oForWO, pVal, _company, sboapp);

                            }
                            else if (pVal.FormUID == "BO_GL" && pVal.ItemUID == "txtCA" && pVal.BeforeAction == false)
                            {
                                SAPbouiCOM.Form oForWO = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                DllProduction.CFLAfterOITMBatchMagnagemet(FormUID, oForWO, pVal, _company, sboapp);

                            }
                            else if (pVal.FormUID == "BO_GL" && pVal.ItemUID == "txtWH" && pVal.BeforeAction == false)
                            {
                                SAPbouiCOM.Form oForWO = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                DllProduction.CFLAfterWareHouse(FormUID, oForWO, pVal, _company, sboapp);

                            }
                            else if (pVal.FormUID == "BO_Par_Production" && pVal.ItemUID == "txtAcct" && pVal.BeforeAction == false)
                            {
                                SAPbouiCOM.Form oForWO = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                DllProduction.CFLAfterOACT(FormUID, oForWO, pVal, _company, sboapp);

                            }
                            else if (pVal.FormUID == "UDO_F_BORP3" && pVal.ItemUID == "0_U_E" && pVal.BeforeAction == false)
                            {
                                SAPbouiCOM.Form oFormNewProductionRoute;
                                oFormNewProductionRoute = sboapp.Forms.Item("UDO_F_BORP3");

                                DllProduction.CFLAfterItemNewProductionRoute(FormUID, oFormNewProductionRoute, pVal, _company, sboapp);

                            }
                            else if (pVal.FormUID == "UDO_F_BORP3" && pVal.ItemUID == "0_U_G" && pVal.BeforeAction == false)
                            {
                                SAPbouiCOM.Form oFormNewProductionRoute;
                                oFormNewProductionRoute = sboapp.Forms.Item("UDO_F_BORP3");

                                DllProduction.CFLAfterOITMMatrix(FormUID, oFormNewProductionRoute, pVal, _company, sboapp);

                            }
                            else if (pVal.FormUID == "BO_Form_EP" && pVal.ItemUID == "txtBP" && pVal.BeforeAction == false)
                            {
                                SAPbouiCOM.Form oFormExternalService;
                                oFormExternalService = sboapp.Forms.Item("BO_Form_EP");

                                DllProduction.CFLAfterBP(FormUID, oFormExternalService, pVal, _company, sboapp);

                            }
                            else if (pVal.FormUID == "BO_Form_EP" && pVal.ItemUID == "txtIC" && pVal.BeforeAction == false)
                            {
                                SAPbouiCOM.Form oFormExternalService;
                                oFormExternalService = sboapp.Forms.Item("BO_Form_EP");

                                DllProduction.CFLAfterOITM(FormUID, oFormExternalService, pVal, _company, sboapp);

                            }
                            else if (pVal.FormUID == "BO_Form_EP" && pVal.ItemUID == "txtWhs" && pVal.BeforeAction == false)
                            {
                                SAPbouiCOM.Form oFormExternalService;
                                oFormExternalService = sboapp.Forms.Item("BO_Form_EP");

                                DllProduction.CFLAfterWhs(FormUID, oFormExternalService, pVal, _company, sboapp);

                            }

                            #endregion
                        }

                        break;

                    case BoEventTypes.et_FORM_DATA_LOAD:

                        if (TieneLicenciaPresupuesto == true)
                        {
                            if (pVal.FormUID == "BOPC" && pVal.FormMode == 3)
                            {
                                #region Eventos_Presupuesto
                                if (Flag1)
                                {
                                    oPresup.presupEventos(BoEventTypes.et_FORM_ACTIVATE, pVal);
                                    FlagNew = false;
                                    #region codigo comentariado
                                    //string usrID;
                                    //Form oForm = null;
                                    //oForm = sboapp.Forms.Item("BOPC");
                                    //EditText et01 = (EditText)oForm.Items.Item("BO_User").Specific;
                                    //EditText et02 = (EditText)oForm.Items.Item("BO_UserID").Specific;
                                    //Funciones.Comunes oFunc = new Funciones.Comunes();
                                    //usrID = oFunc.GetUsrID(_company, sboapp, et01.Value);
                                    //et02.Value = usrID;
                                    //et01.Item.Click();
                                    #endregion
                                    Flag1 = false;
                                }
                                #endregion
                            }
                        }

                        break;

                    case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE:

                        if (TieneLicenciaeBilling == true)
                        {
                            #region Eventos_Presupuesto
                            if (pVal.FormUID == "BOPC")
                            {
                                Form oForm = null;
                                oForm = sboapp.Forms.Item("BOPC");
                                #region codigo comentariado

                                // oForm = sboapp.Forms.ActiveForm;
                                // CheckBox chk01 = (SAPbouiCOM.CheckBox)oForm.Items.Item("BO_ProySN").Specific;
                                // chk01.Checked = true;
                                #endregion
                                oPresup.presupEventos(BoEventTypes.et_FORM_VISIBLE, pVal);
                            }
                            #endregion
                        }

                        break;

                    case SAPbouiCOM.BoEventTypes.et_VALIDATE:

                        if (TieneLicenciaeBilling == true)
                        {
                            #region Validacion_Formulario_eBilling
                            if (pVal.FormUID == "BO_eBillingP" && pVal.FormMode == 3)
                            {
                                SAPbouiCOM.Form oFormVDBO;

                                oFormVDBO = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                            }
                            #endregion
                        }

                        break;

                    case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:

                        if (TieneLicenciaPresupuesto == true)
                        {
                            #region Eventos_Presupuesto
                            if (pVal.FormUID == "BOPC" && pVal.ActionSuccess == true && pVal.BeforeAction == false && FlagNew == true)
                            {
                                oCore.LlenarChkForms(sboapp.Forms.ActiveForm.TypeEx);
                                FlagNew = false;
                            }
                            #endregion
                        }

                        break;

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:

                        #region Eventos_Gestor_AddOn

                        if (pVal.FormUID == "BO_Gestion_AddOn" && pVal.ItemUID == "btnCan" && pVal.BeforeAction == true)

                        {
                            DllFunciones.CloseFormXML(sboapp, "BO_Gestion_AddOn");

                        }
                        else if (pVal.FormUID == "BO_Gestion_AddOn" && pVal.ItemUID == "btnLeft" && pVal.BeforeAction == true)
                        {
                            ActivacionAddIn(sboapp, _company);
                        }
                        else if (pVal.FormUID == "BO_Gestion_AddOn" && pVal.ItemUID == "btnRight" && pVal.BeforeAction == true)
                        {
                            DesactivarAddIn();
                        }

                        else if (pVal.FormUID == "BO_Gestion_AddOn" && pVal.ItemUID == "btnImport" && pVal.BeforeAction == true)
                        {
                            ImportarLicencia();
                        }
                        else if (pVal.FormUID == "BO_Gestion_AddOn" && pVal.ItemUID == "btnVE" && pVal.BeforeAction == true)
                        {
                            DllCore.ValidacionEstructuraAddInsActivos(sboapp, _company, sNameDB);
                        }

                        #endregion                        

                        if (TieneLicenciaPresupuesto == true)
                        {
                            #region Eventos_Presupuesto

                            if (pVal.FormUID == "BOPC" && pVal.FormType == 0 && pVal.ActionSuccess == true && pVal.BeforeAction == false)
                            {
                                oCore.LimpiarChkForm("BOPC", pVal);

                            }

                            #endregion
                        }

                        if (TieneLicenciaeBilling == true)
                        {
                            #region Eventos Facturacion Electronica Emision

                            #region Visor de documentos

                            if (pVal.FormUID == "BOVDEB" && pVal.ItemUID == "btnCan" && pVal.BeforeAction == true)
                            {
                                DllFunciones.CloseFormXML(sboapp, "BOVDEB");
                            }
                            else if (pVal.FormUID == "BOVDEB" && pVal.ItemUID == "btnFind" && pVal.BeforeAction == true)
                            {
                                DlleBilling.InsertDataInMatrix(sboapp, _company);
                            }
                            else if (pVal.FormUID == "BOVDEB" && pVal.ItemUID == "btnVD" && pVal.BeforeAction == true)
                            {
                                SAPbouiCOM.Form oFormVisorDocs = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                                DlleBilling.ActualizarEstadoDocumentos(sboapp, _company, oFormVisorDocs);
                            }
                            else if (pVal.FormUID == "BOVDEB" && pVal.ColUID == "Col_26" && pVal.ItemUID == "MtxOINV" && pVal.BeforeAction == true)
                            {
                                DlleBilling.SincronizarEventosDIAN(sboapp, _company, pVal, pVal.ItemUID, pVal.ColUID);
                            }
                            else if (pVal.FormUID == "BOVDEB" && pVal.ColUID == "Col_28" && pVal.ItemUID == "MtxOINV" && pVal.BeforeAction == true)
                            {
                                DlleBilling.OpenFormEvent(_company, sboapp, pVal);
                            }
                            if (pVal.FormUID == "BOVDEB" && pVal.ItemUID == "btnST" && pVal.BeforeAction == true)
                            {
                                DlleBilling.SelectVRDmatrixlines(sboapp, _company, pVal, pVal.ItemUID, pVal.ColUID);
                            }
                            #endregion

                            #region Visor eventos DIAN 

                            else if (pVal.FormUID == "FormEvDI" && pVal.ItemUID == "btnClose" && pVal.BeforeAction == true)
                            {
                                DllFunciones.CloseFormXML(sboapp, "FormEvDI");

                            }

                            #endregion

                            #region Factura de venta

                            else if (pVal.FormType == 133 && pVal.ItemUID == "FolderBO1" && pVal.Before_Action == true)
                            {
                                #region Adiciona folder eBilling 

                                SAPbouiCOM.Form oFormInvoice;

                                oFormInvoice = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                                DlleBilling.ChangePaneFolderDocuments(oFormInvoice);

                                #endregion
                            }
                            else 
                            if (pVal.FormType == 133 && pVal.ItemUID == "BtnEnvi" && pVal.Before_Action == true)
                            {
                                #region Boton enviar a la DIAN

                                DlleBilling.EnviarDocumentoDIAN(sboapp, _company, pVal.FormType, "ItemEvent", null, "BtnEnvi");
                                
                                #endregion
                            }
                            else if (pVal.FormType == 133 && pVal.ItemUID == "BtnSM" && pVal.Before_Action == true)
                            {
                                #region Abre Formulario Reenviar Email 
                                
                                SAPbouiCOM.Form oFormInvoice = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                sPrefijoDocNumSM = DlleBilling.OpenFormSendMail(sboapp, _company, pVal.FormType, oFormInvoice);

                                DllFunciones.liberarObjetos(oFormInvoice);

                                #endregion
                            }
                            
                            #endregion                            

                            else if (pVal.FormUID == "BO_SM" && pVal.ItemUID == "btnClose" && pVal.Before_Action == true)
                            {
                                #region Cierra Formulario Reenviar E-Mail

                                DllFunciones.CloseFormXML(sboapp, "BO_SM");

                                #endregion
                            }
                            else if (pVal.FormUID == "BO_SM" && pVal.ItemUID == "btnSend" && pVal.Before_Action == true)
                            {
                                DlleBilling.EnviarCorreoElectronico(sboapp, _company, sPrefijoDocNumSM);
                            }

                            #region Factura de reserva de clientes

                            else if (pVal.FormType == 60091 && pVal.ItemUID == "FolderBO1" && pVal.Before_Action == true)
                            {
                                #region Adiciona Folder eBilling 

                                SAPbouiCOM.Form oFormReserveInvoice;

                                oFormReserveInvoice = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                                DlleBilling.ChangePaneFolderDocuments(oFormReserveInvoice);

                                #endregion

                            }
                            else if (pVal.FormType == 60091 && pVal.ItemUID == "BtnEnvi" && pVal.Before_Action == true)
                            {
                                #region Boton enviar a la DIAN

                                DlleBilling.EnviarDocumentoDIAN(sboapp, _company, pVal.FormType, "ItemEvent", null, "BtnEnvi");

                                #endregion

                            }
                            else if (pVal.FormType == 60091 && pVal.ItemUID == "BtnSM" && pVal.Before_Action == true)
                            {
                                #region Abre Formulario Reenviar Email 

                                SAPbouiCOM.Form oFormInvoice = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                sPrefijoDocNumSM = DlleBilling.OpenFormSendMail(sboapp, _company, pVal.FormType, oFormInvoice);

                                DllFunciones.liberarObjetos(oFormInvoice);

                                #endregion
                            }

                            #endregion

                            #region Factura + Pago

                            else if (pVal.FormType == 60090 && pVal.ItemUID == "FolderBO1" && pVal.Before_Action == true)
                            {
                                #region Adiciona pestaña eBilling 

                                SAPbouiCOM.Form oFormPaymentandInvoice;

                                oFormPaymentandInvoice = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                                DlleBilling.ChangePaneFolderDocuments(oFormPaymentandInvoice);

                                #endregion
                            }
                            else if (pVal.FormType == 60090 && pVal.ItemUID == "BtnEnvi" && pVal.Before_Action == true)
                            {
                                #region Boton enviar a la DIAN

                                DlleBilling.EnviarDocumentoDIAN(sboapp, _company, pVal.FormType, "ItemEvent",null, "BtnEnvi");

                                #endregion
                            }
                            else if (pVal.FormType == 60090 && pVal.ItemUID == "BtnSM" && pVal.Before_Action == true)
                            {
                                #region Abre Formulario Reenviar Email 

                                SAPbouiCOM.Form oFormInvoice = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                sPrefijoDocNumSM = DlleBilling.OpenFormSendMail(sboapp, _company, pVal.FormType, oFormInvoice);

                                DllFunciones.liberarObjetos(oFormInvoice);

                                #endregion
                            }

                            #endregion

                            #region Factura de proveedor - Documento Soporte

                            else if (pVal.FormType == 141 && pVal.ItemUID == "FolderBO1" && pVal.Before_Action == true)
                            {
                                #region Adiciona pestaña eBilling 

                                SAPbouiCOM.Form oFormPaymentandInvoice;

                                oFormPaymentandInvoice = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                                DlleBilling.ChangePaneFolderDocuments(oFormPaymentandInvoice);

                                #endregion
                            }
                            else if (pVal.FormType == 141 && pVal.ItemUID == "BtnEnvi" && pVal.Before_Action == true)
                            {
                                #region Boton enviar a la DIAN

                                DlleBilling.EnviarDocumentoDIAN(sboapp, _company, pVal.FormType, "ItemEvent", null, "BtnEnvi");

                                #endregion
                            }

                            #endregion

                            #region Nota Credito Proveedor - Documento Soporte

                            else if (pVal.FormType == 181 && pVal.ItemUID == "FolderBO1" && pVal.Before_Action == true)
                            {
                                #region Adiciona pestaña eBilling 

                                SAPbouiCOM.Form oFormPaymentandInvoice;

                                oFormPaymentandInvoice = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                                DlleBilling.ChangePaneFolderDocuments(oFormPaymentandInvoice);

                                #endregion
                            }
                            else if (pVal.FormType == 181 && pVal.ItemUID == "BtnEnvi" && pVal.Before_Action == true)
                            {
                                #region Boton enviar a la DIAN

                                DlleBilling.EnviarDocumentoDIAN(sboapp, _company, pVal.FormType, "ItemEvent",null, "BtnEnvi");

                                #endregion
                            }

                            #endregion

                            #region Parametros eBiling

                            else if (pVal.FormUID == "BO_eBillingP" && pVal.ItemUID == "btnConTk" && pVal.BeforeAction == true)
                            {
                                SAPbouiCOM.Form oFormParametros = sboapp.Forms.Item("BO_eBillingP");

                                DlleBilling.ConsultaTokens(_company, sboapp, oFormParametros);
                            }
                            else if (pVal.FormUID == "BO_eBillingP" && pVal.ItemUID == "btnVeCam" && pVal.BeforeAction == true)
                            {
                                DlleBilling.CreateUserFieldInvoiceExport(sboapp, _company);
                            }
                            else if (pVal.FormUID == "BO_eBillingP" && pVal.ItemUID == "btnTestAPI" && pVal.BeforeAction == true)
                            {
                                 var response = DlleBilling.FBE_TestAuthorizationAPI(sboapp, _company);
                            }
                            #endregion

                            #region Socios de Negocio

                            else if (pVal.FormType == 134 && pVal.ItemUID == "FolderBO1" && pVal.Before_Action == true)
                            {
                                SAPbouiCOM.Form oFormBP;

                                oFormBP = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                                DlleBilling.ChangePaneFolderBP(oFormBP);
                            }

                            #endregion

                            #region Nota Credito de Clientes

                            else if (pVal.FormType == 179 && pVal.ItemUID == "FolderBO1" && pVal.Before_Action == true)
                            {
                                SAPbouiCOM.Form oFormCreditNote;

                                oFormCreditNote = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                                DlleBilling.ChangePaneFolderDocuments(oFormCreditNote);

                            }
                            else if (pVal.FormType == 179 && pVal.ItemUID == "BtnEnvi" && pVal.Before_Action == true)
                            {
                                #region Boton enviar a la DIAN

                                DlleBilling.EnviarDocumentoDIAN(sboapp, _company, pVal.FormType, "ItemEvent",null, "BtnEnvi");
                                
                                #endregion

                            }
                            else if (pVal.FormType == 179 && pVal.ItemUID == "BtnSM" && pVal.Before_Action == true)
                            {
                                #region Abre Formulario Reenviar Email 

                                SAPbouiCOM.Form oFormInvoice = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                sPrefijoDocNumSM = DlleBilling.OpenFormSendMail(sboapp, _company, pVal.FormType, oFormInvoice);

                                DllFunciones.liberarObjetos(oFormInvoice);

                                #endregion
                            }

                            #endregion

                            #region Nota Debito de clientes

                            else if (pVal.FormType == 65303 && pVal.ItemUID == "FolderBO1" && pVal.Before_Action == true)
                            {
                                SAPbouiCOM.Form oFormCreditNote;

                                oFormCreditNote = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                                DlleBilling.ChangePaneFolderDocuments(oFormCreditNote);

                            }
                            else if (pVal.FormType == 65303 && pVal.ItemUID == "BtnEnvi" && pVal.Before_Action == true)
                            {
                                #region Boton enviar a la DIAN

                                DlleBilling.EnviarDocumentoDIAN(sboapp, _company, pVal.FormType, "ItemEvent",null, "BtnEnvi");

                                #endregion

                            }
                            else if (pVal.FormType == 179 && pVal.ItemUID == "BtnSM" && pVal.Before_Action == true)
                            {
                                #region Abre Formulario Reenviar Email 

                                SAPbouiCOM.Form oFormInvoice = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                sPrefijoDocNumSM = DlleBilling.OpenFormSendMail(sboapp, _company, pVal.FormType, oFormInvoice);

                                DllFunciones.liberarObjetos(oFormInvoice);

                                #endregion
                            }

                            #endregion

                            #endregion
                        }

                        if (TieneLicenciaIntercompany == true)
                        {
                            #region Eventos_Intercompany

                            if (pVal.FormUID == "BO_Visor_Inter" && pVal.ItemUID == "btnCan" && pVal.BeforeAction == true)
                            {
                                DllFunciones.CloseFormXML(sboapp, "BO_Visor_Inter");

                            }
                            else if (pVal.FormType == 1470000200 && pVal.ItemUID == "FolderBO1" && pVal.Before_Action == true)
                            {
                                SAPbouiCOM.Form oFormIntercompany;

                                oFormIntercompany = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                                DllIntercompany.ChangePaneFolder(oFormIntercompany);

                            }

                            else if (pVal.FormType == 142 && pVal.ItemUID == "FolderBO2" && pVal.Before_Action == true)
                            {
                                SAPbouiCOM.Form oFormIntercompany;

                                oFormIntercompany = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                                DllIntercompany.ChangePaneFolder(oFormIntercompany);
                            }

                            #endregion
                        }

                        if (TieneLicenciaProduction == true)
                        {
                            #region Eventos Production

                            #region Ordenes de producción 

                            if (pVal.FormType == 65211 && pVal.ItemUID == "FolderBO1" && pVal.Before_Action == true)
                            {
                                #region Selecciona el Panel "Ruta de produccion"

                                SAPbouiCOM.Form oFormWorkOrder;

                                oFormWorkOrder = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                                DllProduction.ChangePaneFolderWorkOrder(oFormWorkOrder);

                                #endregion

                            }

                            #endregion

                            else if (pVal.FormUID == "BO_Par_Production" && pVal.ItemUID == "btnCan" && pVal.BeforeAction == true)
                            {
                                DllFunciones.CloseFormXML(sboapp, "BO_Par_Production");
                            }
                            else if (pVal.FormUID == "BO_Par_Production" && pVal.ItemUID == "btnUpdate" && pVal.BeforeAction == true)
                            {
                                #region Actualiza la serie de numeracion

                                SAPbouiCOM.Form oFormParProduccion;
                                oFormParProduccion = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                if (oFormParProduccion.Mode == BoFormMode.fm_UPDATE_MODE)
                                {
                                    DllProduction.UpdateParametersProduction(sboapp, _company, oFormParProduccion);
                                }
                                else
                                {
                                    DllFunciones.CloseFormXML(sboapp, "BO_Par_Production");
                                }

                                #endregion

                            }
                            else if (pVal.FormUID == "BOFormCOP" && pVal.ItemUID == "btnCan" && pVal.BeforeAction == true)
                            {
                                DllFunciones.CloseFormXML(sboapp, "BOFormCOP");
                            }
                            else if (pVal.FormUID == "BOFormCOP" && pVal.ItemUID == "btnAct" && pVal.BeforeAction == true)
                            {
                                #region Actualiza el formulario ordenes de produccion

                                DllFunciones.StatusBar(sboapp, BoStatusBarMessageType.smt_Warning, "Consultado ordenes de produccion, por favor espere...");

                                SAPbouiCOM.Form oFormBOPP;
                                oFormBOPP = sboapp.Forms.Item("BOFormCOP");

                                DllProduction.UpdateFormControlProduction(sboapp, _company, oFormBOPP);

                                DllFunciones.StatusBar(sboapp, BoStatusBarMessageType.smt_Success, "Ordenes de produccion consultadas ");

                                #endregion
                            }
                            else if (pVal.FormUID == "BOFormCOP" && pVal.ItemUID == "btnNOP" && pVal.BeforeAction == true)
                            {
                                #region Carga formulario nueva orden de produccion

                                try
                                {
                                    string ArchivoSRF = "New_Work_Order.srf";
                                    DllFunciones.LoadFromXML(sboapp, "BOProduction", ref ArchivoSRF);

                                    SAPbouiCOM.Form oFormNOP;
                                    oFormNOP = sboapp.Forms.Item("BO_New_WO");

                                    DllProduction.LoadFormNewWorkOrder(sboapp, _company, oFormNOP, pVal);

                                }
                                catch (Exception e)
                                {
                                    DllFunciones.sendErrorMessage(sboapp, e);
                                }

                                #endregion
                            }
                            else if (pVal.FormUID == "BOFormCOP" && pVal.ItemUID == "btnNLM" && pVal.BeforeAction == true)
                            {
                                sboapp.Menus.Item("4353").Activate();
                            }
                            else if (pVal.FormUID == "BOFormCOP" && pVal.ItemUID == "btnNSC" && pVal.BeforeAction == true)
                            {
                                sboapp.Menus.Item("39724").Activate();
                            }
                            else if (pVal.FormUID == "BOFormCOP" && pVal.ItemUID == "btnRP" && pVal.Before_Action == true && pVal.Action_Success == false)
                            {
                                #region Carga Formulario Control Ordenes de Produccion

                                try
                                {

                                    string ArchivoSRF = "Production_Route.srf";
                                    DllFunciones.LoadFromXML(sboapp, "BOProduction", ref ArchivoSRF);

                                    SAPbouiCOM.Form oFormBOPR;
                                    oFormBOPR = sboapp.Forms.Item("BO_PR");

                                    DllProduction.ChangueFormProductionRoute(sboapp, _company, oFormBOPR);

                                }
                                catch (Exception e)
                                {

                                    DllFunciones.sendErrorMessage(sboapp, e);
                                }

                                #endregion
                            }
                            else if (pVal.FormUID == "BOFormCOP" && pVal.ItemUID == "btnSearch" && pVal.Before_Action == true && pVal.Action_Success == false)
                            {
                                #region Carga Formulario Control Ordenes de Produccion

                                try
                                {
                                    SAPbouiCOM.Form oFormBOPR;
                                    oFormBOPR = sboapp.Forms.Item("BOFormCOP");

                                    DllProduction.SearchWorkOrder(sboapp, _company, oFormBOPR);

                                }
                                catch (Exception e)
                                {

                                    DllFunciones.sendErrorMessage(sboapp, e);
                                }

                                #endregion
                            }
                            else if (pVal.FormUID == "BOFormMPC" && pVal.ItemUID == "btnSalir" && pVal.BeforeAction == true)
                            {
                                #region Actualiza el formulario materia prima entregada

                                DllFunciones.CloseFormXML(sboapp, "BOFormMPC");

                                #endregion
                            }
                            else if (pVal.FormUID == "BO_New_WO" && pVal.ItemUID == "btnCan" && pVal.BeforeAction == true)
                            {
                                DllFunciones.CloseFormXML(sboapp, "BO_New_WO");
                            }
                            else if (pVal.FormUID == "BO_New_WO" && pVal.ItemUID == "btnNL" && pVal.BeforeAction == true)
                            {
                                #region Adiciona linea en matriz

                                SAPbouiCOM.Form oFormNWO;
                                oFormNWO = sboapp.Forms.Item("BO_New_WO");

                                DllProduction.AddNewRowMatrix(oFormNWO);

                                #endregion

                            }
                            else if (pVal.FormUID == "BO_New_WO" && pVal.ItemUID == "btnDL" && pVal.BeforeAction == true)
                            {
                                #region Se elimina linea de la matrz

                                SAPbouiCOM.Form oFormNWO;
                                oFormNWO = sboapp.Forms.Item("BO_New_WO");

                                DllProduction.DeleteRowMatrix(sboapp, oFormNWO);

                                #endregion

                            }
                            else if (pVal.FormUID == "BO_New_WO" && pVal.ItemUID == "btnNLM" && pVal.BeforeAction == true)
                            {
                                sboapp.Menus.Item("4353").Activate();
                            }
                            else if (pVal.FormUID == "BO_New_WO" && pVal.ItemUID == "btnCreate" && pVal.Before_Action == true && pVal.Action_Success == false)
                            {
                                #region Validacion de campos en Orden de produccion

                                SAPbouiCOM.Form oFormNWO;
                                oFormNWO = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                BubbleEvent = DllProduction.Validate_WorkOrder(sboapp, _company, oFormNWO);

                                if (BubbleEvent == false)
                                {

                                }
                                else
                                {
                                    sCurrentUser = Convert.ToString(_company.UserSignature);

                                    BubbleEvent = DllProduction.Create_Order_Prodcution(sboapp, _company, sMotor, sCurrentUser, oFormNWO);

                                    DllFunciones.CloseFormXML(sboapp, "BO_New_WO");

                                }

                                #endregion
                            }
                            else if (pVal.FormUID == "BO_New_WO" && pVal.ItemUID == "btnAdd" && pVal.Before_Action == true && pVal.Action_Success == false)
                            {
                                #region Adiciona ruta de produccion a la Matrix

                                SAPbouiCOM.Form oFormNWO;
                                oFormNWO = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                DllProduction.AddItemsNWOMatrix(oFormNWO, _company, sboapp);

                                #endregion
                            }
                            else if (pVal.FormUID == "BO_PR" && pVal.ItemUID == "btnNRP" && pVal.Before_Action == true && pVal.Action_Success == false)
                            {
                                #region Abre formulario ruta de produccion

                                string ArchivoSRF = "new_production_route.srf";
                                DllFunciones.LoadFromXML(sboapp, "BOProduction", ref ArchivoSRF);

                                SAPbouiCOM.Form oFormNPR;
                                oFormNPR = sboapp.Forms.Item("UDO_F_BORP3");

                                DllProduction.LoadFormNewProductionRoute(sboapp, _company, oFormNPR, pVal);

                                #endregion
                            }
                            else if (pVal.FormUID == "BO_PR" && pVal.ItemUID == "btnClose" && pVal.Before_Action == true && pVal.Action_Success == false)
                            {
                                #region Cierre el formulario ruta de produccion

                                DllFunciones.CloseFormXML(sboapp, "BO_PR");

                                #endregion
                            }
                            else if (pVal.FormUID == "BO_GL" && pVal.ItemUID == "btnCons" && pVal.Before_Action == true && pVal.Action_Success == false)
                            {
                                #region Consulta los lotes

                                SAPbouiCOM.Form oFormGL;
                                oFormGL = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                DllProduction.SearchBacht(sboapp, _company, oFormGL, pVal);

                                #endregion
                            }
                            else if (pVal.FormUID == "BO_GL" && pVal.ItemUID == "btn1" && pVal.Before_Action == true && pVal.Action_Success == false)
                            {
                                #region Adiciona linea a la matrix LO

                                SAPbouiCOM.Form oFormGL;
                                oFormGL = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                DllProduction.AddLineMatrixLO(oFormGL, pVal);

                                #endregion
                            }
                            else if (pVal.FormUID == "BO_GL" && pVal.ItemUID == "btn2" && pVal.Before_Action == true && pVal.Action_Success == false)
                            {
                                #region Elimina la linea de la matrix LO

                                SAPbouiCOM.Form oFormGL;
                                oFormGL = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                DllProduction.DeleteRowMatrixLO(sboapp, oFormGL, pVal);

                                #endregion
                            }
                            else if (pVal.FormUID == "BO_GL" && pVal.ItemUID == "btn3" && pVal.Before_Action == true && pVal.Action_Success == false)
                            {
                                #region Adiciona la linea de la matrix LD

                                SAPbouiCOM.Form oFormGL;
                                oFormGL = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                DllProduction.AddLineMatrixLD(oFormGL, pVal);

                                #endregion
                            }
                            else if (pVal.FormUID == "BO_GL" && pVal.ItemUID == "btn4" && pVal.Before_Action == true && pVal.Action_Success == false)
                            {
                                #region Adiciona la linea de la matrix LD

                                SAPbouiCOM.Form oFormGL;
                                oFormGL = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                DllProduction.DeleteRowMatrixLD(sboapp, oFormGL, pVal);

                                #endregion
                            }
                            else if (pVal.FormUID == "BO_GL" && pVal.ItemUID == "btnTL" && pVal.Before_Action == true && pVal.Action_Success == false)
                            {
                                BubbleEvent = true;

                                #region Transferencia de Lote

                                SAPbouiCOM.Form oFormGL;
                                oFormGL = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                BubbleEvent = DllProduction.Validate_BachtNumberTrasnfer(sboapp, _company, oFormGL, pVal);

                                if (BubbleEvent == true)
                                {
                                    DllProduction.CreateGoodIssueandGoodReceipt(sboapp, _company, oFormGL);
                                }

                                #endregion
                            }
                            else if (pVal.FormUID == "BO_PR" && pVal.ItemUID == "btnSearch" && pVal.Before_Action == true && pVal.Action_Success == false)
                            {
                                BubbleEvent = true;

                                #region Busqueda de ruta de produccion especifica

                                SAPbouiCOM.Form oFormRP;
                                oFormRP = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                DllProduction.SearchProductionRoute(sboapp, _company, oFormRP);


                                #endregion
                            }
                            else if (pVal.FormUID == "BO_Form_EP" && pVal.ItemUID == "btnCan" && pVal.BeforeAction == true)
                            {
                                DllFunciones.CloseFormXML(sboapp, "BO_Form_EP");
                            }
                            else if (pVal.FormUID == "BO_Form_EP" && pVal.ItemUID == "btnAdd" && pVal.BeforeAction == true)
                            {
                                #region Create Purchase Order

                                int iContinuar = DllFunciones.sendMessageBoxY_N(sboapp, "Se creara la orden de servicio externo, ¿ Desea continuar ?");

                                if (iContinuar == 1)
                                {
                                    SAPbouiCOM.Form oFormPO;
                                    oFormPO = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                    DllProduction.CreatePurchaseOrder(sboapp, _company, oFormPO);

                                }

                                #endregion
                            }
                            else if (pVal.FormUID == "UDO_F_BORP3" && pVal.ItemUID == "btn2" && pVal.Before_Action == true && pVal.Action_Success == false)
                            {
                                #region Elimina la linea de la matrix 

                                SAPbouiCOM.Form oForm;
                                oForm = sboapp.Forms.Item("UDO_F_BORP3");

                                DllProduction.DeleteRowMatrix(sboapp, oForm, pVal, "0_U_G");

                                oForm.Mode = BoFormMode.fm_UPDATE_MODE;

                                #endregion
                            }
                            else if (pVal.FormUID == "UDO_F_BORP3" && pVal.ItemUID == "btn1" && pVal.Before_Action == true && pVal.Action_Success == false)
                            {
                                #region Adiciona la linea de la matrix 

                                SAPbouiCOM.Form oForm;
                                oForm = sboapp.Forms.Item("UDO_F_BORP3");

                                DllProduction.AddLineMatrixRP(oForm, pVal, "0_U_G");

                                if (oForm.Mode == BoFormMode.fm_ADD_MODE)
                                {

                                }
                                else
                                {
                                    oForm.Mode = BoFormMode.fm_UPDATE_MODE;
                                }

                                #endregion
                            }
                            #endregion
                        }

                        if (TieneLicenciaElectronicRepception == true)
                        {
                            #region Eventos Recepcion Electronica

                            #region Visor de documentos

                            if (pVal.FormUID == "BOTVDR" && pVal.ItemUID == "btnDDPT" && pVal.BeforeAction == true)
                            {
                                DllElectronicReception.DownloadDocumentElectronicReceipt(sboapp, _company);
                            }
                            else if (pVal.FormUID == "BOTVDR" && pVal.ItemUID == "btnDAPT" && pVal.BeforeAction == true)
                            {
                                DllElectronicReception.DownloadDocumentXML_PDF(sboapp, _company);
                            }
                            else if (pVal.FormUID == "BOTVDR" && pVal.ItemUID == "cboEvMa" && pVal.BeforeAction == true)
                            {
                                DllElectronicReception.ChagueDocumentStatusBulk(sboapp, _company, pVal, pVal.ColUID);
                            }
                            else if (pVal.FormUID == "BOTVDR" && pVal.ColUID == "Col_17" && pVal.ItemUID == "MtxOPCH" && pVal.BeforeAction == true)
                            {
                                DllElectronicReception.DocumentSearchDIAN(sboapp, _company, pVal, pVal.ColUID);
                            }
                            else if (pVal.FormUID == "BOTVDR" && pVal.ColUID == "Col_20" && pVal.ItemUID == "MtxOPCH" && pVal.BeforeAction == true)
                            {
                                DllElectronicReception.MatrixOpenFile(sboapp, _company, pVal, pVal.ColUID);                                
                            }
                            else if (pVal.FormUID == "BOTVDR" && pVal.ColUID == "Col_21" && pVal.ItemUID == "MtxOPCH" && pVal.BeforeAction == true)
                            {
                                DllElectronicReception.MatrixOpenFile(sboapp, _company, pVal, pVal.ColUID);
                            }
                            else if (pVal.FormUID == "BOTVDR" && pVal.ColUID == "Col_22" && pVal.ItemUID == "MtxOPCH" && pVal.BeforeAction == true)
                            {
                                DllElectronicReception.ChagueStatusDocument(sboapp, _company, pVal, pVal.ColUID);
                            }
                            else if (pVal.FormUID == "BOTVDR" && pVal.ColUID == "Col_7" && pVal.ItemUID == "MtxOPCH" && pVal.BeforeAction == true)
                            {
                                DllElectronicReception.ChagueStatusDocument(sboapp, _company, pVal, pVal.ColUID);
                            }
                            else if (pVal.FormUID == "BOTVDR" && pVal.ColUID == "Col_15" && pVal.ItemUID == "MtxOPCH" && pVal.BeforeAction == true)
                            {
                                DllElectronicReception.ChagueStatusDocument(sboapp, _company, pVal, pVal.ColUID);
                            }
                            else if (pVal.FormUID == "BOTVDR" && pVal.ColUID == "Col_27" && pVal.ItemUID == "MtxOPCH" && pVal.BeforeAction == true)
                            {
                                DllElectronicReception.ChagueStatusDocument(sboapp, _company, pVal, pVal.ColUID);
                            }
                            else if (pVal.FormUID == "BOTVDR" && pVal.ItemUID == "btnCan" && pVal.BeforeAction == true)
                            {
                                DllFunciones.CloseFormXML(sboapp, "BOTVDR");
                            }
                            else if (pVal.FormUID == "BOTVDR" && pVal.ItemUID == "btnFind" && pVal.BeforeAction == true)
                            {
                                SAPbouiCOM.Form oFormVisorDocs = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                                DllElectronicReception.LoadMatrixReception(sboapp, _company, oFormVisorDocs);
                            }
                            else if (pVal.FormUID == "BOTVDR" && pVal.ItemUID == "cbEvMa" && pVal.BeforeAction == true)
                            {                                
                                DllElectronicReception.ChagueDocumentStatusBulk(sboapp, _company, pVal, pVal.ColUID);
                            }
                            else if (pVal.FormUID == "BOTVDR" && pVal.ItemUID == "btnXLS" && pVal.BeforeAction == true)
                            {
                                DllElectronicReception.ExportToExcel(sboapp, _company);
                            }
                            else if (pVal.FormUID == "BO_eBillingP" && pVal.ItemUID == "btnTestAPI" && pVal.BeforeAction == true)
                            {
                                var response = DlleBilling.FBE_TestAuthorizationAPI(sboapp, _company);
                            }

                            #endregion

                            #endregion
                        }

                        break;

                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:

                        if (TieneLicenciaeBilling == true)
                        {
                            #region Eventos eBilling

                            if (pVal.FormType == 133 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.Before_Action == true)
                            {
                                SAPbouiCOM.Form oFormInvoice;

                                oFormInvoice = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.Before_Action == true)
                                {
                                    DlleBilling.AddItemsToDocumets(sboapp, _company, oFormInvoice, "FacturaDeClientes");
                                }
                            }

                            if (pVal.FormType == 141 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.Action_Success == true)
                            {
                                SAPbouiCOM.Form oFormInvoice;

                                oFormInvoice = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.Before_Action == false)
                                {
                                    DlleBilling.AddItemsToDocumets(sboapp, _company, oFormInvoice, "FacturaDeProveedores");
                                }
                            }

                            if (pVal.FormType == 60090 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.Before_Action == true)
                            {
                                SAPbouiCOM.Form oFormPaymentAndInvoice;

                                oFormPaymentAndInvoice = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.Before_Action == true)
                                {
                                    DlleBilling.AddItemsToDocumets(sboapp, _company, oFormPaymentAndInvoice, "FacturaDeClientes");
                                }
                            }

                            if (pVal.FormType == 60091 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.Before_Action == true)
                            {
                                SAPbouiCOM.Form oFormInvoice;

                                oFormInvoice = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.Before_Action == true)
                                {
                                    DlleBilling.AddItemsToDocumets(sboapp, _company, oFormInvoice, "FacturaDeClientes");
                                }
                            }

                            if (pVal.FormType == 179 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.Before_Action == true)
                            {
                                SAPbouiCOM.Form oCreditNote;

                                oCreditNote = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.Before_Action == true)
                                {
                                    DlleBilling.AddItemsToDocumets(sboapp, _company, oCreditNote, "NotaCreditoClientes");
                                }

                            }
                            else if (pVal.FormType == 65303 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.Before_Action == true)
                            {
                                SAPbouiCOM.Form oDebitNote;

                                oDebitNote = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.Before_Action == true)
                                {
                                    DlleBilling.AddItemsToDocumets(sboapp, _company, oDebitNote, "NotaDebitoClientes");
                                }

                            }
                            else if (pVal.FormType == 134 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.Before_Action == true)
                            {
                                SAPbouiCOM.Form oFormBusinessParnerd;

                                oFormBusinessParnerd = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.Before_Action == true)
                                {
                                    DlleBilling.AddItemsToBP(sboapp, _company, oFormBusinessParnerd);
                                }

                            }
                            else if (pVal.FormType == 181 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.Before_Action == true)
                            {
                                SAPbouiCOM.Form oDebitNote;

                                oDebitNote = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.Before_Action == true)
                                {
                                    DlleBilling.AddItemsToDocumets(sboapp, _company, oDebitNote, "NotaCreditoProveedores");
                                }

                            }


                            #endregion
                        }

                        if (TieneLicenciaProduction == true)
                        {
                            #region Produccion Avanzada

                            if (pVal.FormType == 65211 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.Before_Action == true)
                            {
                                SAPbouiCOM.Form oFormWorkOrder;

                                oFormWorkOrder = sboapp.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.Before_Action == true)
                                {
                                    DllProduction.AddItemsToWorkOrder(oFormWorkOrder);
                                }
                            }

                            #endregion
                        }

                        break;

                    case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:

                        if (TieneLicenciaeBilling == true)
                        {
                            #region Eventos Billing

                            if (pVal.FormUID == "BOVDEB" && pVal.ColUID == "Col_9" && pVal.ItemUID == "MtxOINV" && pVal.BeforeAction == true)
                            {
                                SAPbouiCOM.Form oFormVDBO = sboapp.Forms.Item("BOVDEB");

                                DlleBilling.LinkedButtonMatrixFormVDBO(sboapp, _company, oFormVDBO, pVal, "FacturaDeClientes", "Col_9");

                            }
                            else if (pVal.FormUID == "BOVDEB" && pVal.ColUID == "Col_2" && pVal.ItemUID == "MtxOINV" && pVal.BeforeAction == true)
                            {
                                SAPbouiCOM.Form oFormVDBO = sboapp.Forms.Item("BOVDEB");

                                DlleBilling.LinkedButtonMatrixFormVDBO(sboapp, _company, oFormVDBO, pVal, "FacturaDeClientes", "Col_2");

                            }
                            else if (pVal.FormUID == "BOVDEB" && pVal.ColUID == "Col_9" && pVal.ItemUID == "MtxORIN" && pVal.BeforeAction == true)
                            {
                                SAPbouiCOM.Form oFormVDBO = sboapp.Forms.Item("BOVDEB");

                                DlleBilling.LinkedButtonMatrixFormVDBO(sboapp, _company, oFormVDBO, pVal, "NotaCreditoClientes", "Col_9");

                            }
                            else if (pVal.FormUID == "BOVDEB" && pVal.ColUID == "Col_2" && pVal.ItemUID == "MtxORIN" && pVal.BeforeAction == true)
                            {
                                SAPbouiCOM.Form oFormVDBO = sboapp.Forms.Item("BOVDEB");

                                DlleBilling.LinkedButtonMatrixFormVDBO(sboapp, _company, oFormVDBO, pVal, "NotaCreditoClientes", "Col_2");

                            }
                            else if (pVal.FormUID == "BOVDEB" && pVal.ColUID == "Col_9" && pVal.ItemUID == "MtxOINVD" && pVal.BeforeAction == true)
                            {
                                SAPbouiCOM.Form oFormVDBO = sboapp.Forms.Item("BOVDEB");

                                DlleBilling.LinkedButtonMatrixFormVDBO(sboapp, _company, oFormVDBO, pVal, "NotaDebitoClientes", "Col_9");

                            }
                            else if (pVal.FormUID == "BOVDEB" && pVal.ColUID == "Col_2" && pVal.ItemUID == "MtxOINVD" && pVal.BeforeAction == true)
                            {
                                SAPbouiCOM.Form oFormVDBO = sboapp.Forms.Item("BOVDEB");

                                DlleBilling.LinkedButtonMatrixFormVDBO(sboapp, _company, oFormVDBO, pVal, "NotaDebitoClientes", "Col_2");

                            }
                            else if (pVal.FormUID == "BOVDEB" && pVal.ColUID == "Col_9" && pVal.ItemUID == "MtxOPCH" && pVal.BeforeAction == true)
                            {
                                SAPbouiCOM.Form oFormVDBO = sboapp.Forms.Item("BOVDEB");

                                DlleBilling.LinkedButtonMatrixFormVDBO(sboapp, _company, oFormVDBO, pVal, "FacturaDeProveedores", "Col_9");

                            }
                            else if (pVal.FormUID == "BOVDEB" && pVal.ColUID == "Col_2" && pVal.ItemUID == "MtxOPCH" && pVal.BeforeAction == true)
                            {
                                SAPbouiCOM.Form oFormVDBO = sboapp.Forms.Item("BOVDEB");

                                DlleBilling.LinkedButtonMatrixFormVDBO(sboapp, _company, oFormVDBO, pVal, "FacturaDeProveedores", "Col_2");

                            }



                            #endregion
                        }

                        if (TieneLicenciaProduction == true)
                        {
                            #region Eventos Produccion 

                            if (pVal.FormUID == "BOFormCOP" && pVal.ColUID == "Col_0" && pVal.ItemUID == "MtxCOP" && pVal.BeforeAction == true)
                            {
                                SAPbouiCOM.Form oFormCOP = sboapp.Forms.Item("BOFormCOP");

                                DllProduction.LinkedButtonMatrixFormCOP(sboapp, _company, oFormCOP, pVal, "Col_0");

                            }
                            if (pVal.FormUID == "BOFormCOP" && pVal.ColUID == "Col_13" && pVal.ItemUID == "MtxCOP" && pVal.BeforeAction == true)
                            {
                                SAPbouiCOM.Form oFormCOP = sboapp.Forms.Item("BOFormCOP");

                                DllProduction.LinkedButtonMatrixFormCOP(sboapp, _company, oFormCOP, pVal, "Col_13");

                            }
                            else if (pVal.FormUID == "BOFormCOP" && pVal.ColUID == "Col_2" && pVal.ItemUID == "MtxCOP" && pVal.BeforeAction == true)
                            {
                                SAPbouiCOM.Form oFormCOP = sboapp.Forms.Item("BOFormCOP");

                                DllProduction.LinkedButtonMatrixFormCOP(sboapp, _company, oFormCOP, pVal, "Col_1");

                            }
                            else if (pVal.FormUID == "BOFormCOP" && pVal.ColUID == "Col_8" && pVal.ItemUID == "MtxCOP" && pVal.BeforeAction == true)
                            {
                                SAPbouiCOM.Form oFormCOP = sboapp.Forms.Item("BOFormCOP");

                                DllProduction.LinkedButtonMatrixFormCOP(sboapp, _company, oFormCOP, pVal, "Col_8");

                            }
                            else if (pVal.FormUID == "BOFormMPC" && pVal.ColUID == "Col_1" && pVal.ItemUID == "MtxMPE" && pVal.BeforeAction == true)
                            {
                                SAPbouiCOM.Form oFormCOP = sboapp.Forms.Item("BOFormMPC");

                                DllProduction.LinkedButtonMatrixFormMPE(sboapp, _company, oFormCOP, pVal, "Col_1");

                            }


                            #endregion
                        }


                        break;

                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:

                        if (TieneLicenciaProduction == true)
                        {
                            #region Eventos Produccion 

                            if (pVal.FormUID == "BO_GL" && pVal.ColUID == "Col_0" && pVal.ItemUID == "mtxLO" && pVal.ActionSuccess == true)
                            {
                                SAPbouiCOM.Form oFormGL = sboapp.Forms.Item("BO_GL");

                                DllProduction.CantidadesExistentes(sboapp, _company, oFormGL, pVal);

                            }

                            #endregion
                        }

                        break;
                }

            }
            catch (Exception e)
            {
                DllFunciones.sendErrorMessage(sboapp, e);
            }

        }

        private void Sboapp_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            Presupuesto.Core oCore = new Presupuesto.Core(sboapp, _company);
            BubbleEvent = true;

            if (!pVal.BeforeAction)
            {
                return;
            }

            Form frm = null;

            switch (pVal.MenuUID)
            {

                case "mnuBO_ConfPresup":

                    #region Eventos en el Menu Presupuesto

                    frm = sboapp.Forms.ActiveForm;

                    try
                    {
                        //si el formulario ya estaba creado le hace focus
                        SAPbouiCOM.Form oFormPresu = sboapp.Forms.Item("BO_ConfPresup");
                        oFormPresu.Select();
                    }
                    catch
                    {
                        //Si el formulario no existe lo crea
                        Presupuesto.Core oFormPresu = new Presupuesto.Core(sboapp, _company);
                        oFormPresu.showForm("BO_ConfPresup");
                    }


                    #endregion

                    break;

                case "mnuBO_PerfilPresup":

                    #region Eventos en el Menu Presupuesto

                    frm = sboapp.Forms.ActiveForm;

                    try
                    {
                        //si el formulario ya estaba creado le hace focus
                        SAPbouiCOM.Form oFormPerfilPresup = sboapp.Forms.Item("BO_PerfilPresup");
                        oFormPerfilPresup.Select();
                    }
                    catch
                    {
                        //Si el formulario no existe lo crea
                        Presupuesto.Core oFormPerfilPresup = new Presupuesto.Core(sboapp, _company);
                        oFormPerfilPresup.showForm("BO_PerfilPresup");
                    }


                    #endregion

                    break;

                case "mnuBO_PresupCuenta":

                    #region Eventos en el Menu Presupuesto

                    frm = sboapp.Forms.ActiveForm;

                    try
                    {
                        //si el formulario ya estaba creado le hace focus
                        SAPbouiCOM.Form oFormPresupCuenta = sboapp.Forms.Item("BO_PresupCuenta");
                        oFormPresupCuenta.Select();
                    }
                    catch
                    {
                        //Si el formulario no existe lo crea
                        Presupuesto.Core oFormPresupCuenta = new Presupuesto.Core(sboapp, _company);
                        oFormPresupCuenta.showForm("BO_PresupCuenta");
                    }

                    #endregion

                    break;

                case "1282":

                    #region Eventos en el Menu Presupuesto

                    if (sboapp.Forms.ActiveForm.TypeEx == "BOPC")
                    {
                        Console.WriteLine("Click en nuevo");
                        FlagNew = true;
                        oCore.LlenarChkForms(sboapp.Forms.ActiveForm.TypeEx);

                    }

                    #endregion

                    #region Eventos en el Menu eBilling

                    if (TieneLicenciaeBilling == true)
                    {
                        if (sboapp.Forms.ActiveForm.TypeEx == "133" && pVal.BeforeAction == true)
                        {
                            frm = sboapp.Forms.ActiveForm;

                            DlleBilling.ItemsLabelStatusDIAN(frm, "133", "MenuEvent");
                        }
                    }

                    #endregion

                    break;

                case "1281":

                    #region Eventos en el Menu Presupuesto

                    if (sboapp.Forms.ActiveForm.TypeEx == "BOPC")
                    {
                        Console.WriteLine("Click en buscar");
                    }

                    #endregion

                    #region Eventos en el Menu eBilling

                    if (TieneLicenciaeBilling == true)
                    {
                        if (sboapp.Forms.ActiveForm.TypeEx == "133" && pVal.BeforeAction == true)
                        {
                            frm = sboapp.Forms.ActiveForm;

                            DlleBilling.ItemsLabelStatusDIAN(frm, "133", "MenuEvent");
                        }
                    }

                    #endregion

                    break;

                case "1289":


                    break;

                case "1304":


                    break;
                case "mnuBO_GestorAddOn":

                    #region Eventos en el Menu Gestor AddOn                   

                    try
                    {
                        DllCore.OpenFormGestorAddOn(sboapp, _company);

                    }
                    catch (Exception e)
                    {

                        DllFunciones.sendErrorMessage(sboapp, e);

                    }                    

                    #endregion

                    break;

                case "mnuBO_VD":

                    #region EVentos en el Menu Intercompany   

                    if (TieneLicenciaIntercompany == true)
                    {
                        #region Abre formulario visor de documentos intercompany

                        try
                        {
                            string ArchivoSRF = "Visor_Documentos.srf";
                            DllFunciones.LoadFromXML(sboapp, "Intercompany", ref ArchivoSRF);
                        }
                        catch (Exception)
                        {

                            throw;
                        }

                        #endregion
                    }

                    #endregion

                    break;

                case "smBO_eBil01":

                    #region Menu Parametros Iniciales FE

                    if (TieneLicenciaeBilling == true)
                    {
                        #region Abre fomrulario parametros Facturacion Electronica 

                        try
                        {
                            DlleBilling.OpenFormParametrosIniciales(sboapp, _company, sMotor, Convert.ToString(_company.UserSignature));

                        }
                        catch (Exception e)
                        {

                            DllFunciones.sendErrorMessage(sboapp, e);

                        }

                        #endregion
                    }

                    #endregion

                    break;

                case "smHXRE01":

                    #region Eventos en el Menu eBilling 

                    if (TieneLicenciaElectronicRepception == true)
                    {
                        #region Abre fomrulario parametros Facturacion Electronica 

                        try
                        {
                            DllElectronicReception.OpenFormParametrosIniciales(sboapp, _company, sMotor);
                        }
                        catch (Exception e)
                        {

                            DllFunciones.sendErrorMessage(sboapp, e);

                        }

                        #endregion
                    }

                    #endregion

                    break;
                case "smBO_eBil02":

                    #region Abre formulario unidades de medida 

                    if (TieneLicenciaeBilling == true)
                    {
                        try
                        {
                            DlleBilling.OpenFormUnidadesMedida(sboapp, _company, sMotor);

                        }
                        catch (Exception e)
                        {

                            DllFunciones.sendErrorMessage(sboapp, e);

                        }

                    }

                    #endregion

                    break;

                case "smBO_eBil03":

                    #region Abre Formulario visor de documentos

                    if (TieneLicenciaeBilling == true)
                    {
                        try
                        {
                            DlleBilling.OpenFormVisorDocumentosEnviados(sboapp, _company, sMotor);
                        }
                        catch (Exception e)
                        {
                            DllFunciones.sendErrorMessage(sboapp, e);
                        }

                    }

                    #endregion

                    break;

                case "AddRowMtx":

                    #region Abre matrix para adicionar series de numeracion en formulario de para parametizacion

                    if (TieneLicenciaeBilling == true)
                    {

                        try
                        {
                            SAPbouiCOM.Form oFormBO_eBillingP = sboapp.Forms.Item("BO_eBillingP");

                            DlleBilling.AddRowMatrix(oFormBO_eBillingP);
                        }
                        catch (Exception)
                        {

                            throw;
                        }


                    }

                    #endregion

                    break;

                case "AddOSM":

                    if (TieneLicenciaProduction == true)
                    {

                        string ArchivoSRF = "External_Service.srf";
                        DllFunciones.LoadFromXML(sboapp, "BOElectronicReception", ref ArchivoSRF);

                        SAPbouiCOM.Form oFormExternalService;
                        oFormExternalService = sboapp.Forms.Item("BOTVDR");

                        DllProduction.oFormExternalSource(sboapp, _company, oFormExternalService);
                    }

                    break;

                case "mnuBO_PP":

                    #region Eventos en el Menu Produccion Avanzada

                    if (TieneLicenciaProduction == true)
                    {
                        #region Carga formulario parametros produccion

                        try
                        {
                            string ArchivoSRF = "Parameters_production.srf";
                            DllFunciones.LoadFromXML(sboapp, "BOProduction", ref ArchivoSRF);

                            SAPbouiCOM.Form oFormBOPP;
                            oFormBOPP = sboapp.Forms.Item("BO_Par_Production");

                            DllProduction.LoadFormParProduction(sboapp, _company, oFormBOPP);

                        }
                        catch (Exception e)
                        {
                            DllFunciones.sendErrorMessage(sboapp, e);
                        }

                        #endregion
                    }

                    #endregion

                    break;

                case "mnuBO_COP":

                    #region Carga Formulario Control Ordenes de Produccion

                    if (TieneLicenciaProduction == true)
                    {
                        try
                        {
                            DllFunciones.StatusBar(sboapp, BoStatusBarMessageType.smt_Warning, "Consultado ordenes de produccion, por favor espere...");

                            string ArchivoSRF = "Control_Production.srf";
                            DllFunciones.LoadFromXML(sboapp, "BOProduction", ref ArchivoSRF);

                            SAPbouiCOM.Form oFormBOPP;
                            oFormBOPP = sboapp.Forms.Item("BOFormCOP");

                            DllProduction.ChangueFormControlProduction(sboapp, _company, oFormBOPP);

                            DllFunciones.StatusBar(sboapp, BoStatusBarMessageType.smt_Success, "Ordenes de produccion consultadas ");

                        }
                        catch (Exception e)
                        {

                            DllFunciones.sendErrorMessage(sboapp, e);
                        }



                    }

                    #endregion

                    break;

                case "mnuBO_RP":

                    #region Carga Formulario Control Ordenes de Produccion

                    if (TieneLicenciaProduction == true)
                    {

                        try
                        {

                            string ArchivoSRF = "Production_Route.srf";
                            DllFunciones.LoadFromXML(sboapp, "BOProduction", ref ArchivoSRF);

                            SAPbouiCOM.Form oFormBOPR;
                            oFormBOPR = sboapp.Forms.Item("BO_PR");

                            DllProduction.ChangueFormProductionRoute(sboapp, _company, oFormBOPR);

                        }
                        catch (Exception e)
                        {

                            DllFunciones.sendErrorMessage(sboapp, e);
                        }

                    }

                    #endregion

                    break;

                case "mnuBO_GL":

                    #region Carga Formulario Control Ordenes de Produccion

                    if (TieneLicenciaProduction == true)
                    {

                        try
                        {

                            string ArchivoSRF = "batch_management.srf";
                            DllFunciones.LoadFromXML(sboapp, "BOProduction", ref ArchivoSRF);

                            SAPbouiCOM.Form oFormGL;
                            oFormGL = sboapp.Forms.Item("BO_GL");

                            DllProduction.LoadFormBatchManagement(sboapp, _company, oFormGL);

                        }
                        catch (Exception e)
                        {

                            DllFunciones.sendErrorMessage(sboapp, e);
                        }

                    }

                    #endregion

                    break;

                case "smBO_eBil04":

                    #region Carga Formulario Visor de documentos Reception

                    if (TieneLicenciaElectronicRepception == true)
                    {

                        try
                        {
                            DllElectronicReception.OpenFormVisorDocumentosRecibidos(sboapp, _company, sMotor);

                        }
                        catch (Exception e)
                        {

                            DllFunciones.sendErrorMessage(sboapp, e);

                        }

                    }

                    #endregion

                    break;

            }

        }

        private void Sboapp_RightClick(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            #region Eliminacion de Menus

            if (sboapp.Menus.Exists("AddRowMtx"))
            {
                sboapp.Menus.RemoveEx("AddRowMtx");
            }

            if (sboapp.Menus.Exists("AddSM"))
            {
                sboapp.Menus.RemoveEx("AddSM");
            }

            if (sboapp.Menus.Exists("AddOSM"))
            {
                sboapp.Menus.RemoveEx("AddOSM");
            }

            #endregion

            try
            {
                SAPbouiCOM.Form oFormActive = sboapp.Forms.ActiveForm;

                if (oFormActive.UniqueID == "BO_eBillingP" && eventInfo.BeforeAction == true && oFormActive.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    DlleBilling.Right_Click(ref eventInfo, sboapp, "1");
                }
                else if (oFormActive.TypeEx == "133" && eventInfo.BeforeAction == true && oFormActive.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    DlleBilling.Right_Click(ref eventInfo, sboapp, "2");
                }

                else if (oFormActive.TypeEx == "65211" && eventInfo.BeforeAction == true && oFormActive.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {

                    DllProduction.Right_Click(ref eventInfo, sboapp);

                }


            }
            catch (Exception)
            {

                throw;
            }
        }

        private void Sboapp_DataEvent(ref SAPbouiCOM.BusinessObjectInfo ByRef, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                switch (ByRef.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:

                        if (TieneLicenciaeBilling)
                        {
                            #region Factura de venta

                            if (ByRef.ActionSuccess == true && ByRef.FormTypeEx == "133")
                            {
                                #region Actualiza Label Status DIAN

                                SAPbouiCOM.Form frm = sboapp.Forms.Item(ByRef.FormUID);

                                DlleBilling.ItemsLabelStatusDIAN(frm, ByRef.FormTypeEx.ToString(), "DataEvent");

                                #endregion
                            }

                            #endregion

                            #region Factura + pago

                            else if (ByRef.ActionSuccess == true && ByRef.FormTypeEx == "60090")
                            {
                                #region Actualiza Label Status DIAN

                                SAPbouiCOM.Form frm = sboapp.Forms.Item(ByRef.FormUID);

                                DlleBilling.ItemsLabelStatusDIAN(frm, ByRef.FormTypeEx.ToString(), "DataEvent");

                                #endregion

                            }

                            #endregion

                            #region Factura reserva

                            else if (ByRef.ActionSuccess == true && ByRef.FormTypeEx == "60091")
                            {
                                #region Actualiza Label Status DIAN

                                SAPbouiCOM.Form frm = sboapp.Forms.Item(ByRef.FormUID);

                                DlleBilling.ItemsLabelStatusDIAN(frm, ByRef.FormTypeEx.ToString(), "DataEvent");



                                #endregion

                            }

                            #endregion

                            #region Nota credito de clientes

                            else if (ByRef.ActionSuccess == true && ByRef.FormTypeEx == "179")
                            {
                                #region Actualiza Label Status DIAN

                                SAPbouiCOM.Form frm = sboapp.Forms.Item(ByRef.FormUID);

                                DlleBilling.ItemsLabelStatusDIAN(frm, ByRef.FormTypeEx.ToString(), "DataEvent");

                                #endregion

                            }

                            #endregion

                            #region Factura de compra

                            else if (ByRef.ActionSuccess == true && ByRef.FormTypeEx == "141")
                            {
                                #region Actualiza Label Status DIAN

                                SAPbouiCOM.Form frm = sboapp.Forms.Item(ByRef.FormUID);

                                DlleBilling.ItemsLabelStatusDIAN(frm, ByRef.FormTypeEx.ToString(), "DataEvent");

                                #endregion

                            }

                            #endregion

                            #region Parametros Iniciales

                            else if (ByRef.ActionSuccess == true && ByRef.FormTypeEx == "UDO_FT_BO_eBillingP")
                            {
                                #region Actualiza Label Status DIAN

                                DlleBilling.CheckBoxCloudControlCenterAndExportInvoice(sboapp, _company);

                                #endregion
                            }

                            #endregion
                        }

                        break;

                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:

                        if (TieneLicenciaeBilling)
                        {

                            #region Factura de venta 

                            if (ByRef.ActionSuccess == true && ByRef.FormTypeEx == "133")
                            {
                                SAPbouiCOM.Form frm = sboapp.Forms.Item(ByRef.FormUID);

                                #region Boton enviar a la DIAN

                                DlleBilling.EnviarDocumentoDIAN(sboapp, _company, 133, "DataEvent", ByRef, "DataEvent") ;

                                #endregion

                                DlleBilling.ItemsLabelStatusDIAN(frm, ByRef.FormTypeEx.ToString(), "DataEvent");
                            }

                            #endregion

                            #region Factura + pago
                            
                            else if (ByRef.ActionSuccess == true && ByRef.FormTypeEx == "60090")
                            {
                                SAPbouiCOM.Form frm = sboapp.Forms.Item(ByRef.FormUID);

                                #region Boton enviar a la DIAN

                                DlleBilling.EnviarDocumentoDIAN(sboapp, _company, 60090, "ItemEvent", ByRef, "DataEvent");

                                #endregion

                                DlleBilling.ItemsLabelStatusDIAN(frm, ByRef.FormTypeEx.ToString(), "DataEvent");

                            }

                            #endregion

                            #region Factura de reserva 

                            else if (ByRef.ActionSuccess == true && ByRef.FormTypeEx == "60091")
                            {
                                #region Boton enviar a la DIAN

                                DlleBilling.EnviarDocumentoDIAN(sboapp, _company, 60091, "DataEvent", ByRef, "DataEvent");

                                #endregion

                            }

                            #endregion

                            #region Nota credito de clientes

                            else if (ByRef.ActionSuccess == true && ByRef.FormTypeEx == "179")
                            {
                                #region Boton enviar a la DIAN

                                DlleBilling.EnviarDocumentoDIAN(sboapp, _company, 179, "DataEvent", ByRef, "DataEvent");

                                #endregion

                            }

                            #endregion

                            #region Factura de proveedor

                            if (ByRef.ActionSuccess == true && ByRef.FormTypeEx == "141")
                            {
                                SAPbouiCOM.Form frm = sboapp.Forms.Item(ByRef.FormUID);

                                #region Boton enviar a la DIAN

                                DlleBilling.EnviarDocumentoDIAN(sboapp, _company, 141, "DataEvent", ByRef, "DataEvent");

                                #endregion

                                DlleBilling.ItemsLabelStatusDIAN(frm, ByRef.FormTypeEx.ToString(), "DataEvent");
                            }

                            #endregion


                        }

                        break;

                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:

                        if (TieneLicenciaeBilling)
                        {
                            if (ByRef.ActionSuccess == true && ByRef.FormTypeEx == "UDO_FT_BO_eBillingP")
                            {
                                DlleBilling.UpdateChangueProveedorTecnologico(sboapp,_company);
                            }
                        }

                        break;

                }

            }
            catch (Exception)
            {

            }
        }

        #region Metodos

        public void AdicionSubmenu(SAPbouiCOM.MenuCreationParams _oCreationPackage, SAPbouiCOM.Menus _oMenus, SAPbouiCOM.MenuItem _oMenuItem, Application _sboapp, string AddIn, string _sNameDB)
        {

            try
            {
                #region Variables y Objetos

                string sQuerieValidacion = null;
                string LicenciaAddIn = null;
                string ParametrosLicencia = null;
                string LicenciaValida = null;

                XmlDocument XmlConsulta = XmlConsulta = new XmlDocument();

                sPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                sMotor = Convert.ToString(_company.DbServerType);

                SAPbobsCOM.Recordset oValidacion = ((SAPbobsCOM.Recordset)_company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));

                #endregion

                #region Instanciacion Dll

                BOBillingConsolidator.BillingConsolidator DllBillingConsolidator = new BillingConsolidator();

                #endregion

                sQuerieValidacion = DllFunciones.GetStringXMLDocument(_company, "Core", "ValidacionAddOnBO", "ValidacionAddIn");
                sQuerieValidacion = sQuerieValidacion.Replace("%AddIn%", AddIn);

                oValidacion.DoQuery(sQuerieValidacion);

                LicenciaAddIn = oValidacion.Fields.Item(0).Value.ToString();

                if (string.IsNullOrEmpty(LicenciaAddIn))
                {
                    if (AddIn == "AddInPresupuesto")
                    {
                        TieneLicenciaPresupuesto = false;
                    }

                    if (AddIn == "AddInIntercompany")
                    {
                        TieneLicenciaIntercompany = false;
                    }

                    if (AddIn == "AddIneBillingBO")
                    {
                        TieneLicenciaeBilling = false;
                    }

                    if (AddIn == "AddInProduccion")
                    {
                        TieneLicenciaProduction = false;
                    }

                    if (AddIn == "AddInElectronicReception")
                    {
                        TieneLicenciaElectronicRepception = false;
                    }

                    if (AddIn == "AddInBillingConsolidator")
                    {
                        TieneLicenciaConsolidadorFacturacion = false;
                    }

                }
                else
                {
                    ParametrosLicencia = _sboapp.Company.InstallationId + "_" + _sboapp.Company.ServerName + "_" + AddIn;
                    LicenciaValida = DllCore.Encriptar(ParametrosLicencia, Llave);

                    if (LicenciaAddIn == LicenciaValida && AddIn == "AddInPresupuesto")
                    {
                        #region Menu Presupuesto

                        _oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                        _oCreationPackage.UniqueID = "mnuBO_Presupuesto";
                        _oCreationPackage.Enabled = true;
                        _oCreationPackage.String = "Presupuesto Avanzado";
                        _oCreationPackage.Image = "";
                        _oCreationPackage.Position = _oCreationPackage.Position + 1;
                        _oMenus = _oMenuItem.SubMenus;
                        _oMenus.AddEx(_oCreationPackage);

                        // Sub-sub menu Parametrizaciones
                        _oMenus = _oMenuItem.SubMenus;
                        _oMenuItem = sboapp.Menus.Item("mnuBO_Presupuesto");
                        _oCreationPackage.Enabled = true;
                        _oCreationPackage.Position = _oCreationPackage.Position + 1;
                        _oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                        _oCreationPackage.UniqueID = "mnuBO_ParamPresup";
                        _oCreationPackage.String = "Parametrizaciones";
                        _oCreationPackage.Image = "";
                        _oMenus = _oMenuItem.SubMenus;
                        _oMenuItem.SubMenus.AddEx(_oCreationPackage);

                        _oMenuItem = sboapp.Menus.Item("mnuBO_ParamPresup");
                        _oCreationPackage.Enabled = true;
                        _oCreationPackage.Position = _oCreationPackage.Position + 1;
                        _oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        _oCreationPackage.UniqueID = "mnuBO_ConfPresup";
                        _oCreationPackage.String = "Configuración";
                        _oCreationPackage.Image = "";
                        _oMenuItem.SubMenus.AddEx(_oCreationPackage);

                        _oMenuItem = sboapp.Menus.Item("mnuBO_ParamPresup");
                        _oCreationPackage.Enabled = true;
                        _oCreationPackage.Position = _oCreationPackage.Position + 1;
                        _oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        _oCreationPackage.UniqueID = "mnuBO_PerfilPresup";
                        _oCreationPackage.String = "Perfiles";
                        _oCreationPackage.Image = "";
                        _oMenuItem.SubMenus.AddEx(_oCreationPackage);

                        _oMenuItem = sboapp.Menus.Item("mnuBO_Presupuesto");
                        _oCreationPackage.Enabled = true;
                        _oCreationPackage.Position = _oCreationPackage.Position + 1;
                        _oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        _oCreationPackage.UniqueID = "mnuBO_PresupCuenta";
                        _oCreationPackage.String = "Presupuesto por Cuenta";
                        _oCreationPackage.Image = "";
                        _oMenuItem.SubMenus.AddEx(_oCreationPackage);

                        _oMenuItem = sboapp.Menus.Item("mnuBO_Presupuesto");
                        _oCreationPackage.Enabled = true;
                        _oCreationPackage.Position = _oCreationPackage.Position + 1;
                        _oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        _oCreationPackage.UniqueID = "mnuBO_PresupVendor";
                        _oCreationPackage.String = "Presupuesto por Vendedor";
                        _oCreationPackage.Image = "";
                        _oMenuItem.SubMenus.AddEx(_oCreationPackage);

                        _oMenuItem = sboapp.Menus.Item("mnuBO_Presupuesto");
                        _oCreationPackage.Enabled = true;
                        _oCreationPackage.Position = _oCreationPackage.Position + 1;
                        _oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        _oCreationPackage.UniqueID = "mnuBO_PresupItem";
                        _oCreationPackage.String = "Presupuesto por Artículo";
                        _oCreationPackage.Image = "";
                        _oMenuItem.SubMenus.AddEx(_oCreationPackage);

                        _oMenuItem = sboapp.Menus.Item("mnuBO_Presupuesto");
                        _oCreationPackage.Enabled = true;
                        _oCreationPackage.Position = _oCreationPackage.Position + 1;
                        _oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        _oCreationPackage.UniqueID = "mnuBO_PresupCliente";
                        _oCreationPackage.String = "Presupuesto por Cliente";
                        _oCreationPackage.Image = "";
                        _oMenuItem.SubMenus.AddEx(_oCreationPackage);

                        _oMenuItem = sboapp.Menus.Item("mnuBO_Presupuesto");
                        _oCreationPackage.Enabled = true;
                        _oCreationPackage.Position = _oCreationPackage.Position + 1;
                        _oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        _oCreationPackage.UniqueID = "mnuBO_PresupProv";
                        _oCreationPackage.String = "Presupuesto por Proveedor";
                        _oCreationPackage.Image = "";
                        _oMenuItem.SubMenus.AddEx(_oCreationPackage);

                        _oMenuItem = sboapp.Menus.Item("mnuBO_Presupuesto");
                        _oCreationPackage.Enabled = true;
                        _oCreationPackage.Position = _oCreationPackage.Position + 1;
                        _oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        _oCreationPackage.UniqueID = "mnuBO_PresupProyec";
                        _oCreationPackage.String = "Presupuesto por Proyecto";
                        _oCreationPackage.Image = "";
                        _oMenuItem.SubMenus.AddEx(_oCreationPackage);

                        _oMenuItem = sboapp.Menus.Item("mnuBO_Presupuesto");
                        _oCreationPackage.Enabled = true;
                        _oCreationPackage.Position = _oCreationPackage.Position + 1;
                        _oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        _oCreationPackage.UniqueID = "mnuBO_PresupDim";
                        _oCreationPackage.String = "Presupuesto por Dimensión";
                        _oCreationPackage.Image = "";
                        _oMenuItem.SubMenus.AddEx(_oCreationPackage);

                        _oMenus = _oMenuItem.SubMenus;
                        _oMenuItem = sboapp.Menus.Item("mnuBO_Presupuesto");
                        _oCreationPackage.Enabled = true;
                        _oCreationPackage.Position = _oCreationPackage.Position + 1;
                        _oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                        _oCreationPackage.UniqueID = "mnuBO_InfoPresup";
                        _oCreationPackage.String = "Informes de Presupuesto";
                        _oCreationPackage.Image = "";
                        _oMenus = _oMenuItem.SubMenus;
                        _oMenuItem.SubMenus.AddEx(_oCreationPackage);
                        string sXML = null;
                        sXML = sboapp.Menus.GetAsXML();
                        System.Xml.XmlDocument xmlD = new System.Xml.XmlDocument();
                        xmlD.LoadXml(sXML);
                        //xmlD.Save("c:\\mnu.xml");
                        xmlD.Save((Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Core\\Menus\\mnu.xml"));

                        TieneLicenciaPresupuesto = true;

                        #endregion
                    }
                    else if (LicenciaAddIn == LicenciaValida && AddIn == "AddInIntercompany")
                    {
                        #region Menu Intercompañia

                        _oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                        _oCreationPackage.UniqueID = "mnuBO_Inter";
                        _oCreationPackage.Enabled = true;
                        _oCreationPackage.String = "Intercompañia";
                        _oCreationPackage.Image = "";
                        _oCreationPackage.Position = _oCreationPackage.Position + 1;
                        _oMenus = _oMenuItem.SubMenus;
                        _oMenus.AddEx(_oCreationPackage);

                        _oMenuItem = sboapp.Menus.Item("mnuBO_Inter");
                        _oCreationPackage.Enabled = true;
                        _oCreationPackage.Position = _oCreationPackage.Position + 1;
                        _oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        _oCreationPackage.UniqueID = "mnuBO_VD";
                        _oCreationPackage.String = "Visor Documentos";
                        _oCreationPackage.Image = "";
                        _oMenuItem.SubMenus.AddEx(_oCreationPackage);

                        TieneLicenciaIntercompany = true;

                        #endregion
                    }
                    else if (LicenciaAddIn == LicenciaValida && AddIn == "AddIneBillingBO")
                    {
                        #region Menu Facturacion Electronica - Emision

                        _oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                        _oCreationPackage.UniqueID = "mnuBO_eBil01";
                        _oCreationPackage.Enabled = true;
                        _oCreationPackage.String = "Facturación Electronica";
                        _oCreationPackage.Image = "";
                        _oCreationPackage.Position = _oCreationPackage.Position + 1;
                        _oMenus = _oMenuItem.SubMenus;
                        _oMenus.AddEx(_oCreationPackage);

                        _oMenuItem = sboapp.Menus.Item("mnuBO_eBil01");
                        _oCreationPackage.Enabled = true;
                        _oCreationPackage.Position = _oCreationPackage.Position + 1;
                        _oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        _oCreationPackage.UniqueID = "smBO_eBil01";
                        _oCreationPackage.String = "Parametros iniciales";
                        _oCreationPackage.Image = "";
                        _oMenuItem.SubMenus.AddEx(_oCreationPackage);

                        _oMenuItem = sboapp.Menus.Item("mnuBO_eBil01");
                        _oCreationPackage.Enabled = true;
                        _oCreationPackage.Position = _oCreationPackage.Position + 1;
                        _oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        _oCreationPackage.UniqueID = "smBO_eBil02";
                        _oCreationPackage.String = "Unidades de Medida";
                        _oCreationPackage.Image = "";
                        _oMenuItem.SubMenus.AddEx(_oCreationPackage);

                        _oMenuItem = sboapp.Menus.Item("mnuBO_eBil01");
                        _oCreationPackage.Enabled = true;
                        _oCreationPackage.Position = _oCreationPackage.Position + 1;
                        _oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        _oCreationPackage.UniqueID = "smBO_eBil03";
                        _oCreationPackage.String = "Visor de documentos - Enviados";
                        _oCreationPackage.Image = "";
                        _oMenuItem.SubMenus.AddEx(_oCreationPackage);

                        TieneLicenciaeBilling = true;

                        #endregion
                    }
                    else if (LicenciaAddIn == LicenciaValida && AddIn == "AddInElectronicReception")
                    {
                        #region Menu Facturacion Electronica - Recepcion

                        if (sboapp.Menus.Exists("mnuBO_eBil01"))
                        {
                            _oMenuItem = sboapp.Menus.Item("mnuBO_eBil01");
                            _oCreationPackage.Enabled = true;
                            _oCreationPackage.Position = _oCreationPackage.Position + 1;
                            _oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                            _oCreationPackage.UniqueID = "smBO_eBil04";
                            _oCreationPackage.String = "Visor de documentos - Recibidos";
                            _oCreationPackage.Image = "";
                            _oMenuItem.SubMenus.AddEx(_oCreationPackage);
                        }
                        else
                        {
                            _oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                            _oCreationPackage.UniqueID = "mnuBO_eBil01";
                            _oCreationPackage.Enabled = true;
                            _oCreationPackage.String = "Facturación Electronica";
                            _oCreationPackage.Image = "";
                            _oCreationPackage.Position = _oCreationPackage.Position + 1;
                            _oMenus = _oMenuItem.SubMenus;
                            _oMenus.AddEx(_oCreationPackage);

                            _oMenuItem = sboapp.Menus.Item("mnuBO_eBil01");
                            _oCreationPackage.Enabled = true;
                            _oCreationPackage.Position = _oCreationPackage.Position + 1;
                            _oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                            _oCreationPackage.UniqueID = "smHXRE01"; 
                            _oCreationPackage.String = "Parametros iniciales";
                            _oCreationPackage.Image = "";
                            _oMenuItem.SubMenus.AddEx(_oCreationPackage);

                            _oMenuItem = sboapp.Menus.Item("mnuBO_eBil01");
                            _oCreationPackage.Enabled = true;
                            _oCreationPackage.Position = _oCreationPackage.Position + 1;
                            _oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                            _oCreationPackage.UniqueID = "smBO_eBil04";
                            _oCreationPackage.String = "Visor de documentos - Recibidos";
                            _oCreationPackage.Image = "";
                            _oMenuItem.SubMenus.AddEx(_oCreationPackage);
                        }

                        TieneLicenciaElectronicRepception = true;

                        #endregion
                    }
                    else if (LicenciaAddIn == LicenciaValida && AddIn == "AddInProduccion")
                    {
                        #region Menu Produccion Avanzada

                        _oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                        _oCreationPackage.UniqueID = "mnuBO_P";
                        _oCreationPackage.Enabled = true;
                        _oCreationPackage.String = "Producción Avanzada";
                        _oCreationPackage.Image = "";
                        _oCreationPackage.Position = _oCreationPackage.Position + 1;
                        _oMenus = _oMenuItem.SubMenus;
                        _oMenus.AddEx(_oCreationPackage);

                        _oMenuItem = sboapp.Menus.Item("mnuBO_P");
                        _oCreationPackage.Enabled = true;
                        _oCreationPackage.Position = _oCreationPackage.Position + 1;
                        _oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        _oCreationPackage.UniqueID = "mnuBO_PP";
                        _oCreationPackage.String = "Parametos iniciales";
                        _oCreationPackage.Image = "";
                        _oMenuItem.SubMenus.AddEx(_oCreationPackage);

                        _oMenuItem = sboapp.Menus.Item("mnuBO_P");
                        _oCreationPackage.Enabled = true;
                        _oCreationPackage.Position = _oCreationPackage.Position + 1;
                        _oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        _oCreationPackage.UniqueID = "mnuBO_RP";
                        _oCreationPackage.String = "Rutas de producción";
                        _oCreationPackage.Image = "";
                        _oMenuItem.SubMenus.AddEx(_oCreationPackage);

                        _oMenuItem = sboapp.Menus.Item("mnuBO_P");
                        _oCreationPackage.Enabled = true;
                        _oCreationPackage.Position = _oCreationPackage.Position + 1;
                        _oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        _oCreationPackage.UniqueID = "mnuBO_COP";
                        _oCreationPackage.String = "Ordenes de produccion";
                        _oCreationPackage.Image = "";
                        _oMenuItem.SubMenus.AddEx(_oCreationPackage);

                        _oMenuItem = sboapp.Menus.Item("mnuBO_P");
                        _oCreationPackage.Enabled = true;
                        _oCreationPackage.Position = _oCreationPackage.Position + 1;
                        _oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        _oCreationPackage.UniqueID = "mnuBO_GL";
                        _oCreationPackage.String = "Gestión de lotes";
                        _oCreationPackage.Image = "";
                        _oMenuItem.SubMenus.AddEx(_oCreationPackage);

                        TieneLicenciaProduction = true;

                        #endregion
                    }
                    else if (LicenciaAddIn == LicenciaValida && AddIn == "AddInBillingConsolidator")
                    {
                        #region Menu Produccion Avanzada

                        DllBillingConsolidator.AddMenu_BillingConsolidator(_oCreationPackage, _oMenus, _oMenuItem, _sboapp, _company, AddIn, _sNameDB);

                        TieneLicenciaConsolidadorFacturacion = true;

                        #endregion
                    }
                    else
                    {
                        DllFunciones.sendMessageBox(sboapp, "El " + AddIn + " no tiene una licencia valida para el AddOn, por favor comunicarse con el administrador del sistema - B-One Tech..");
                    }
                }
                DllFunciones.liberarObjetos(oValidacion);
            }

            catch (Exception)
            {

                throw;
            }
        }

        public void ActivacionAddIn(Application _sboapp, SAPbobsCOM.Company _oCompany)
        {
            try
            {
                #region Variables y objetos

                SAPbouiCOM.Form oBO_Gestion_AddOn;
                SAPbobsCOM.UserTable oUDTAdminAddOn;
                SAPbouiCOM.Grid oGridAddInDisponibles;
                SAPbobsCOM.Recordset oChkStatusAddIn = null;

                string sChkStatusAddIn = null;
                string sLicencia = null;
                string sLocalizacion = null;
                string AddInSeleccionado = null;

                int Rsd = 0;

                #endregion

                #region Intanciacion Objetos

                oBO_Gestion_AddOn = sboapp.Forms.Item("BO_Gestion_AddOn");

                oGridAddInDisponibles = (SAPbouiCOM.Grid)(oBO_Gestion_AddOn.Items.Item("GridDispo").Specific);
                oChkStatusAddIn = ((SAPbobsCOM.Recordset)_company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));

                #endregion

                #region Consulta Localizacion 

                sLocalizacion = ((SAPbouiCOM.ComboBox)(oBO_Gestion_AddOn.Items.Item("txtLoca").Specific)).Value.ToString();

                #endregion

                #region Consulta AddIn Seleccionado 

                AddInSeleccionado = DllFunciones.GetFieldGridFromSelectedRow(oGridAddInDisponibles, "AddIn");

                sChkStatusAddIn = DllFunciones.GetStringXMLDocument(_company, "Core", "ValidacionAddOnBO", "CheckAddInActivo");
                sChkStatusAddIn = sChkStatusAddIn.Replace("%AddInSeleccionado%", AddInSeleccionado);

                oChkStatusAddIn.DoQuery(sChkStatusAddIn);
                sChkStatusAddIn = oChkStatusAddIn.Fields.Item(0).Value.ToString();

                #endregion


                if (sChkStatusAddIn == "S" || sChkStatusAddIn == "A")
                {
                    DllFunciones.sendMessageBox(sboapp, "El Add-In seleccionado ya fue activado para la empresa");
                }
                else
                {
                    if (string.IsNullOrEmpty(AddInSeleccionado))
                    {

                        DllFunciones.sendMessageBox(sboapp, "No ha seleccionado ningun Add-In para activar, por favor revise");

                    }
                    else if (string.IsNullOrEmpty(sLocalizacion))
                    {

                        DllFunciones.sendMessageBox(sboapp, "No ha seleccionado el tipo de localización, por favor revise");

                    }
                    else
                    {

                        sLicencia = oChkStatusAddIn.Fields.Item(1).Value.ToString();
                        oUDTAdminAddOn = (SAPbobsCOM.UserTable)(_company.UserTables.Item("BOAdminAddOn"));

                        if (string.IsNullOrEmpty(sLicencia))
                        {

                            #region Si no tiene licencia actualiza El AddIn sin licencia

                            oUDTAdminAddOn.GetByKey(AddInSeleccionado);
                            oUDTAdminAddOn.UserFields.Fields.Item("U_Status").Value = "S";

                            Rsd = oUDTAdminAddOn.Update();

                            #endregion

                        }
                        else
                        {

                            #region Si tiene licencia actualiza el AddIn con Licencia

                            oUDTAdminAddOn.GetByKey(AddInSeleccionado);
                            oUDTAdminAddOn.UserFields.Fields.Item("U_Status").Value = "A";
                            Rsd = oUDTAdminAddOn.Update();

                            #endregion

                        }

                        if (Rsd == 0)
                        {
                            #region Si se actualiza el campo licencia libera los objetos

                            DllFunciones.liberarObjetos(oUDTAdminAddOn);
                            DllFunciones.liberarObjetos(oChkStatusAddIn);
                            DllFunciones.liberarObjetos(oGridAddInDisponibles);

                            #endregion

                            #region Se procede a crear los campos del AddIn Seleccionado

                            DllFunciones.sendMessageBox(sboapp, "Se crearan los campos necesarios para el funcionamiento del Add-In " + AddInSeleccionado + ", Por favor asegurar que todos los usuarios este fuera del sistema");

                            if (AddInSeleccionado == "AddInIntercompany")
                            {

                                DllIntercompany.CreacionTablasyCamposeIntercompany();
                            }
                            else if (AddInSeleccionado == "AddInPresupuesto")
                            {

                            }
                            else if (AddInSeleccionado == "AddIneBillingBO")
                            {

                                DlleBilling.CreacionTablasyCamposeBillingBO(sboapp, _company, sMotor, sLocalizacion);

                            }
                            else if (AddInSeleccionado == "AddInProduccion")
                            {

                                DllProduction.CreateUDTandUDFProduction(sboapp, _company);

                            }
                            else if (AddInSeleccionado == "AddInElectronicReception")
                            {
                                DllElectronicReception.CreacionTablasyCamposeBillingBO(sboapp, _company);
                            }
                            else if (AddInSeleccionado == "AddInElectronicReception")
                            {                                

                            }

                            DllFunciones.sendMessageBox(sboapp, "AddIn Activado correctamente, por favor salir y volver a ingresar a SAP");

                            #region Cierra Formulario

                            DllCore.CloseFormGestorAddOn(_sboapp, _oCompany);

                            #endregion
                            

                            oBO_Gestion_AddOn.Close();

                            #endregion

                        }
                        else
                        {

                            string Mensaje_error = null;
                            Mensaje_error = Convert.ToString(_company.GetLastErrorCode()) + " : " + _company.GetLastErrorDescription();

                            DllFunciones.sendMessageBox(sboapp, Mensaje_error);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                DllFunciones.sendErrorMessage(sboapp, e);
            }
        }

        public void DesactivarAddIn()
        {
            try
            {
                SAPbouiCOM.Form oBO_Gestion_AddOn;
                SAPbobsCOM.UserTable oUDTAdminAddOn;
                SAPbouiCOM.Grid oGridAddInActivo;
                SAPbobsCOM.Recordset oChkStatusAddIn = null;
                string sChkStatusAddIn = null;

                int Rsd = 0;
                XmlDocument XmlConsulta;
                XmlConsulta = new XmlDocument();

                oBO_Gestion_AddOn = sboapp.Forms.Item("BO_Gestion_AddOn");

                oGridAddInActivo = (SAPbouiCOM.Grid)(oBO_Gestion_AddOn.Items.Item("GridActi").Specific);
                oChkStatusAddIn = ((SAPbobsCOM.Recordset)_company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));

                sPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                sMotor = Convert.ToString(_company.DbServerType);

                string AddInSeleccionado = DllFunciones.GetFieldGridFromSelectedRow(oGridAddInActivo, "AddIn");

                switch (sMotor)
                {
                    case "dst_MSSQL2012":
                    case "dst_MSSQL2014":
                    case "dst_MSSQL2016":
                    case "dst_MSSQL2017":

                        XmlConsulta.Load(sPath + "\\Core\\Queries\\ValidacionAddOnBOSQL.xml");
                        sChkStatusAddIn = XmlConsulta.SelectSingleNode("Queries/CheckAddInActivo").InnerText;
                        sChkStatusAddIn = sChkStatusAddIn.Replace("%AddInSeleccionado%", AddInSeleccionado);

                        break;

                    case "dst_HANADB":

                        XmlConsulta.Load(sPath + "\\Core\\Queries\\ValidacionAddOnBOHANA.xml");
                        sChkStatusAddIn = XmlConsulta.SelectSingleNode("Queries/CheckAddInActivo").InnerText;
                        sChkStatusAddIn = sChkStatusAddIn.Replace("%AddInSeleccionado%", AddInSeleccionado);

                        break;

                    default:
                        break;
                }

                oChkStatusAddIn.DoQuery(sChkStatusAddIn);
                sChkStatusAddIn = oChkStatusAddIn.Fields.Item(0).Value.ToString();

                if (sChkStatusAddIn == "I")
                {
                    DllFunciones.sendMessageBox(sboapp, "El Add-In seleccionado ya fue desactivado para la empresa");
                    oBO_Gestion_AddOn.Close();
                }
                else
                {
                    if (string.IsNullOrEmpty(AddInSeleccionado))
                    {
                        DllFunciones.sendMessageBox(sboapp, "No ha seleccionado ningun Add-In para desactivar, por favor revise");
                    }
                    else
                    {
                        oUDTAdminAddOn = (SAPbobsCOM.UserTable)(_company.UserTables.Item("BOAdminAddOn"));
                        oUDTAdminAddOn.GetByKey(AddInSeleccionado);
                        oUDTAdminAddOn.UserFields.Fields.Item("U_Status").Value = "I";

                        Rsd = oUDTAdminAddOn.Update();

                        if (Rsd == 0)
                        {
                            DllFunciones.sendMessageBox(sboapp, "Add-In desactivado correctamente, por favor salir y volver a ingresar a SAP");
                            oBO_Gestion_AddOn.Close();
                        }
                        else
                        {
                            string Mensaje_error = null;
                            Mensaje_error = Convert.ToString(_company.GetLastErrorCode()) + " : " + _company.GetLastErrorDescription();

                            DllFunciones.sendMessageBox(sboapp, Mensaje_error);
                        }
                    }
                }
                DllFunciones.liberarObjetos(oChkStatusAddIn);
            }
            catch (Exception e)
            {
                DllFunciones.sendErrorMessage(sboapp, e);
            }
        }

        public void ImportarLicencia()
        {
            try
            {
                SAPbouiCOM.Form oBO_Gestion_AddOn;
                SAPbobsCOM.UserTable oUDTAdminAddOn = null;
                SAPbouiCOM.EditText Licencia = null;
                //SAPbobsCOM.Recordset oChKLicense = null;
                SAPbouiCOM.Grid oGridAddInActives = null;
                //string sChKLicense = null;
                string AddInActive = null;
                XmlDocument XmlConsulta;
                int Rsd = 0;

                oBO_Gestion_AddOn = sboapp.Forms.Item("BO_Gestion_AddOn");
                Licencia = ((SAPbouiCOM.EditText)(oBO_Gestion_AddOn.Items.Item("txtLicen").Specific));
                oGridAddInActives = (SAPbouiCOM.Grid)(oBO_Gestion_AddOn.Items.Item("GridActi").Specific);

                XmlConsulta = new XmlDocument();

                AddInActive = DllFunciones.GetFieldGridFromSelectedRow(oGridAddInActives, "AddIn");

                if (string.IsNullOrEmpty(AddInActive))
                {
                    DllFunciones.sendMessageBox(sboapp, "Por favor seleccione el Add-In al que desea asignarle la licencia la licencia");
                }
                else
                {
                    if (string.IsNullOrEmpty(Licencia.Value))
                    {
                        DllFunciones.sendMessageBox(sboapp, "Por favor ingresar una licencia para poder realizar la importacion");
                    }
                    else
                    {
                        oUDTAdminAddOn = (SAPbobsCOM.UserTable)(_company.UserTables.Item("BOAdminAddOn"));
                        oUDTAdminAddOn.GetByKey(AddInActive);
                        oUDTAdminAddOn.UserFields.Fields.Item("U_Licencia").Value = Licencia.Value;
                        oUDTAdminAddOn.UserFields.Fields.Item("U_Status").Value = "A";

                        Rsd = oUDTAdminAddOn.Update();

                        if (Rsd == 0)
                        {
                            DllFunciones.sendMessageBox(sboapp, "Licencia importada correctamente, por favor salir y volver a ingresar a SAP");
                        }
                        else
                        {
                            string Mensaje_error = null;
                            Mensaje_error = Convert.ToString(_company.GetLastErrorCode()) + " : " + _company.GetLastErrorDescription();
                            DllFunciones.sendMessageBox(sboapp, Mensaje_error);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                DllFunciones.sendErrorMessage(sboapp, e);
            }
        }

        public static string Encriptar(string texto, string Llave)
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

        public static string Desencriptar(string textoEncriptado, string Llave)
        {
            try
            {
                byte[] keyArray;
                byte[] Array_a_Descifrar = Convert.FromBase64String(textoEncriptado);

                //algoritmo MD5
                MD5CryptoServiceProvider hashmd5 = new MD5CryptoServiceProvider();

                keyArray = hashmd5.ComputeHash(UTF8Encoding.UTF8.GetBytes(Llave));

                hashmd5.Clear();

                TripleDESCryptoServiceProvider tdes = new TripleDESCryptoServiceProvider();

                tdes.Key = keyArray;
                tdes.Mode = CipherMode.ECB;
                tdes.Padding = PaddingMode.PKCS7;

                ICryptoTransform cTransform = tdes.CreateDecryptor();

                byte[] resultArray = cTransform.TransformFinalBlock(Array_a_Descifrar, 0, Array_a_Descifrar.Length);

                tdes.Clear();
                textoEncriptado = UTF8Encoding.UTF8.GetString(resultArray);

            }
            catch (Exception)
            {

            }
            return textoEncriptado;
        }
        

        #endregion
    }

}
