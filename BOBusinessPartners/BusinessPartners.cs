using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using Funciones;


namespace BOBusinessPartners
{
    public class BusinessPartners
    {
        public void CreacionClientesCimpa(SAPbobsCOM.Company _oCompany, string _PathFileLog)
        {
            #region Instanciacion 

            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            #endregion

            try
            {
                #region Variables y Objetos

                SAPbobsCOM.Recordset oConsultaClientesCRM = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oInsertSocioNegociosCRM = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                SAPbobsCOM.Recordset oConsultaDireccionesCRM = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oInsertDireccionesCRM = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                SAPbobsCOM.Recordset oConsultaCotizacionesCabeceraCRM = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oConsultaCotizacionesLineasCRM = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                SAPbobsCOM.Recordset oUpdateCotizacionesCRM = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                #endregion

                #region Obtiene consultas

                string sGetSociosNegocioCRM = DllFunciones.GetStringXMLDocument(_oCompany, "BOBusinessPartners", "BusinessPartners", "GetBusinessPartnersCRM");
                string sGetDireccionesCRM = DllFunciones.GetStringXMLDocument(_oCompany, "BOBusinessPartners", "BusinessPartners", "GetDireccionesCRM");
                string sGetCotizacionesCabeceraCRM = DllFunciones.GetStringXMLDocument(_oCompany, "BOBusinessPartners", "BusinessPartners", "GetCotizacionesCabeceraCRM");
                string sGetCotizacionesLineasCRM = DllFunciones.GetStringXMLDocument(_oCompany, "BOBusinessPartners", "BusinessPartners", "GetCotizacionesLineasCRM");
                string sOriginalGetCotizacionesLineasCRM = sGetCotizacionesLineasCRM;

                string sGetUpdateBusinessPartnersCRMOriginal = DllFunciones.GetStringXMLDocument(_oCompany, "BOBusinessPartners", "BusinessPartners", "GetInsertBusinessPartnersCRM");
                string sGetUpdateDireccionesCRMOriginal = DllFunciones.GetStringXMLDocument(_oCompany, "BOBusinessPartners", "BusinessPartners", "GetUpdateDireccionesCRM");
                string sGetUpdateCotizacionesCRMOriginal = DllFunciones.GetStringXMLDocument(_oCompany, "BOBusinessPartners", "BusinessPartners", "GetUpdateCotizacionesCRM");

                string sGetUpdateBusinessPartnersCRMError = DllFunciones.GetStringXMLDocument(_oCompany, "BOBusinessPartners", "BusinessPartners", "GetInsertBusinessPartnersCRMError");
                string sGetUpdateCotizacionesCRMError = DllFunciones.GetStringXMLDocument(_oCompany, "BOBusinessPartners", "BusinessPartners", "GetUpdateCotizacionesCRMError");
                string sOriginalGetUpdateCotizacionesCRMError = DllFunciones.GetStringXMLDocument(_oCompany, "BOBusinessPartners", "BusinessPartners", "GetUpdateCotizacionesCRMError");


                #endregion

                #region Creacion datos Maestros Lead

                DllFunciones.Logger("Ejecuando consulta [BO_spSincronizacionSocioNegociosCRM]", _PathFileLog);

                oConsultaClientesCRM.DoQuery(sGetSociosNegocioCRM);

                if (oConsultaClientesCRM.RecordCount > 0)
                {
                    #region Insercion Lead en SAP

                    SAPbobsCOM.BusinessPartners oBPLead = (SAPbobsCOM.BusinessPartners)_oCompany.GetBusinessObject(BoObjectTypes.oBusinessPartners);

                    oConsultaClientesCRM.MoveFirst();

                    do
                    {
                        DllFunciones.Logger("Creando Lead : " + Convert.ToString(oConsultaClientesCRM.Fields.Item("U_CP_CardCode").Value.ToString()), _PathFileLog);

                        DllFunciones.Logger(Convert.ToString(oConsultaClientesCRM.Fields.Item("U_CP_CardCode").Value.ToString()), _PathFileLog);
                        oBPLead.CardCode = Convert.ToString(oConsultaClientesCRM.Fields.Item("U_CP_CardCode").Value.ToString());
                        oBPLead.CardName = Convert.ToString(oConsultaClientesCRM.Fields.Item("U_CP_CardName").Value.ToString());
                        oBPLead.CardType = BoCardTypes.cLid;
                        oBPLead.GroupCode = Convert.ToInt32(oConsultaClientesCRM.Fields.Item("U_CP_GroupCode").Value.ToString());
                        oBPLead.Currency = Convert.ToString(oConsultaClientesCRM.Fields.Item("U_CP_Currency").Value.ToString());
                        oBPLead.PayTermsGrpCode = Convert.ToInt32(oConsultaClientesCRM.Fields.Item("U_CP_GroupNum").Value.ToString());
                        oBPLead.Phone1 = Convert.ToString(oConsultaClientesCRM.Fields.Item("U_CP_Phone1").Value.ToString());
                        oBPLead.Cellular = Convert.ToString(oConsultaClientesCRM.Fields.Item("U_CP_Cellular").Value.ToString());
                        oBPLead.MailAddress = Convert.ToString(oConsultaClientesCRM.Fields.Item("U_CP_E_Mail").Value.ToString());
                        oBPLead.FederalTaxID = Convert.ToString(oConsultaClientesCRM.Fields.Item("U_CP_FederalTaxId").Value.ToString());
                        oBPLead.SalesPersonCode = Convert.ToInt32(oConsultaClientesCRM.Fields.Item("U_CP_SlpCode").Value.ToString());
                        oBPLead.Notes = Convert.ToString(oConsultaClientesCRM.Fields.Item("U_CP_CardForeignName").Value.ToString());
                        int Rsd = oBPLead.Add();

                        if (Rsd == 0)
                        {
                            string sGetUpdateBusinessPartnersCRM = sGetUpdateBusinessPartnersCRMOriginal.Replace("%CardCode%", Convert.ToString(oConsultaClientesCRM.Fields.Item("U_CP_CardCode").Value.ToString()));
                            DllFunciones.Logger("Lead Creado Existosamente", _PathFileLog);
                            oInsertSocioNegociosCRM.DoQuery(sGetUpdateBusinessPartnersCRM);
                        }
                        else
                        {

                            if (_oCompany.GetLastErrorCode() == 1320000140)
                            {
                                sGetUpdateBusinessPartnersCRMError = sGetUpdateBusinessPartnersCRMError.Replace("%CardCode%", Convert.ToString(oConsultaClientesCRM.Fields.Item("U_CP_CardCode").Value.ToString())).Replace("%Error%", "Cliente Lead ya existe");

                                //DllFunciones.Logger(sGetUpdateBusinessPartnersCRMError, _PathFileLog);

                                oInsertSocioNegociosCRM.DoQuery(sGetUpdateBusinessPartnersCRMError);
                            }
                            else
                            {
                                sGetUpdateBusinessPartnersCRMError = sGetUpdateBusinessPartnersCRMError.Replace("%CardCode%", Convert.ToString(oConsultaClientesCRM.Fields.Item("U_CP_CardCode").Value.ToString())).Replace("%Error%", "Cliente Lead ya existe");

                                //DllFunciones.Logger(sGetUpdateBusinessPartnersCRMError, _PathFileLog);

                                oInsertSocioNegociosCRM.DoQuery(sGetUpdateBusinessPartnersCRMError);
                            }                            
                        }

                        oConsultaClientesCRM.MoveNext();

                    } while (oConsultaClientesCRM.EoF == false);

                    #endregion
                }
                else
                {
                    DllFunciones.Logger("No se encotnraron socios led a integrar", _PathFileLog);
                }

                #endregion

                #region Creacion Direcciones Lead datos Maestros Lead

                DllFunciones.Logger("Ejecuando consulta [BO_spSincronizacionDireccionesCRM]", _PathFileLog);

                oConsultaDireccionesCRM.DoQuery(sGetDireccionesCRM);

                if (oConsultaDireccionesCRM.RecordCount > 0)
                {
                    #region Insercion direcciones en SAP

                    oConsultaDireccionesCRM.MoveFirst();

                    int iSetCurrentLine = 0;

                    SAPbobsCOM.BusinessPartners oBPLeadDirecciones = (SAPbobsCOM.BusinessPartners)_oCompany.GetBusinessObject(BoObjectTypes.oBusinessPartners);

                    do
                    {

                        DllFunciones.Logger("Creando Direccion Lead : " + Convert.ToString(oConsultaDireccionesCRM.Fields.Item("U_CP_BPCode").Value.ToString()), _PathFileLog);
                        oBPLeadDirecciones.GetByKey(Convert.ToString(oConsultaDireccionesCRM.Fields.Item("U_CP_BPCode").Value.ToString()));

                        oBPLeadDirecciones.Addresses.SetCurrentLine(iSetCurrentLine);
                        oBPLeadDirecciones.Addresses.AddressName = Convert.ToString(oConsultaDireccionesCRM.Fields.Item("U_CP_AddressName").Value.ToString());
                        oBPLeadDirecciones.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_ShipTo;
                        oBPLeadDirecciones.Addresses.Block = Convert.ToString(oConsultaDireccionesCRM.Fields.Item("U_CP_Block").Value.ToString());
                        oBPLeadDirecciones.Addresses.Street = Convert.ToString(oConsultaDireccionesCRM.Fields.Item("U_CP_Street").Value.ToString());
                        oBPLeadDirecciones.Addresses.ZipCode = Convert.ToString(oConsultaDireccionesCRM.Fields.Item("U_CP_ZipCode").Value.ToString());
                        oBPLeadDirecciones.Addresses.City = Convert.ToString(oConsultaDireccionesCRM.Fields.Item("U_CP_City").Value.ToString());
                        oBPLeadDirecciones.Addresses.Country = Convert.ToString(oConsultaDireccionesCRM.Fields.Item("U_CP_Country").Value.ToString());

                        int Rsd = oBPLeadDirecciones.Update();

                        if (Rsd == 0)
                        {
                            string sGetUpdateDireccionesCRM = sGetUpdateDireccionesCRMOriginal.Replace("%CardCode%", Convert.ToString(oConsultaDireccionesCRM.Fields.Item("U_CP_BPCode").Value.ToString()));

                            oInsertDireccionesCRM.DoQuery(sGetUpdateDireccionesCRM);

                        }
                        else
                        {
                            string sGetUpdateDireccionesCRM = sGetUpdateDireccionesCRMOriginal.Replace("%CardCode%", Convert.ToString(oConsultaDireccionesCRM.Fields.Item("U_CP_BPCode").Value.ToString())).Replace("%Error%", "No se agregar la direccion");

                            oInsertDireccionesCRM.DoQuery(sGetUpdateDireccionesCRM);
                        }

                        oConsultaDireccionesCRM.MoveNext();

                    } while (oConsultaDireccionesCRM.EoF == false);

                    #endregion
                }
                else
                {
                    DllFunciones.Logger("No se encotraron direcciones a integrar", _PathFileLog);

                }

                #endregion

                #region Creacion Cotizaciones 

                DllFunciones.Logger("Ejecuando consulta [BO_spSincronizacionCotizacionesCabeceraCRM]", _PathFileLog);

                oConsultaCotizacionesCabeceraCRM.DoQuery(sGetCotizacionesCabeceraCRM);                

                if (oConsultaCotizacionesCabeceraCRM.RecordCount > 0)
                {
                    oConsultaCotizacionesCabeceraCRM.MoveFirst();

                    do
                    {
                        #region Insercion Cotizaciones en SAP                   

                        #region Valida si existen Lineas Asociadas a la cotizacion

                        sGetCotizacionesLineasCRM = sOriginalGetCotizacionesLineasCRM.Replace("%DocEntry%", Convert.ToString(oConsultaCotizacionesCabeceraCRM.Fields.Item("U_CP_HXC").Value.ToString()));

                        DllFunciones.Logger(sGetCotizacionesLineasCRM, _PathFileLog);

                        oConsultaCotizacionesLineasCRM.DoQuery(sGetCotizacionesLineasCRM);

                        #endregion

                        if (oConsultaCotizacionesLineasCRM.RecordCount > 0)
                        {
                            #region Inserta Pedido

                            SAPbobsCOM.Documents oQuotation = (SAPbobsCOM.Documents)_oCompany.GetBusinessObject(BoObjectTypes.oQuotations);

                            DllFunciones.Logger("Creando Cotizacion : " + Convert.ToString(oConsultaCotizacionesCabeceraCRM.Fields.Item("U_CP_HXC").Value.ToString()), _PathFileLog);

                            oQuotation.CardCode = Convert.ToString(oConsultaCotizacionesCabeceraCRM.Fields.Item("U_CP_CardCode").Value.ToString());
                            oQuotation.DocDate = DateTime.Now;
                            oQuotation.DocDueDate = DateTime.Now;
                            oQuotation.NumAtCard = Convert.ToString(oConsultaCotizacionesCabeceraCRM.Fields.Item("U_CP_HXC").Value.ToString());

                            #region Adiciona Lineas

                            oConsultaCotizacionesLineasCRM.MoveFirst();
                            do
                                {
                                    oQuotation.Lines.ItemCode = Convert.ToString(oConsultaCotizacionesLineasCRM.Fields.Item("U_CP_ItemCode").Value.ToString());
                                    oQuotation.Lines.Quantity = Convert.ToDouble(oConsultaCotizacionesLineasCRM.Fields.Item("U_CP_Quantity").Value.ToString());
                                    oQuotation.Lines.UnitPrice = Convert.ToDouble(oConsultaCotizacionesLineasCRM.Fields.Item("U_CP_PriceBefDisc").Value.ToString());
                                    oQuotation.Lines.WarehouseCode = Convert.ToString(oConsultaCotizacionesLineasCRM.Fields.Item("U_CP_WhsCode").Value.ToString());

                                    oQuotation.Lines.Add();

                                    oConsultaCotizacionesLineasCRM.MoveNext();

                                } while (oConsultaCotizacionesLineasCRM.EoF == false);

                                #endregion

                            oQuotation.Comments = Convert.ToString(oConsultaCotizacionesCabeceraCRM.Fields.Item("U_CP_Header").Value.ToString());

                            DllFunciones.Logger("Final", _PathFileLog);

                            int Rsd = oQuotation.Add();

                            if (Rsd == 0)
                            {
                                string sGetUpdateCotizacionesCRM = sGetUpdateCotizacionesCRMOriginal.Replace("%DocEntry%", Convert.ToString(oConsultaCotizacionesCabeceraCRM.Fields.Item("U_CP_HXC").Value.ToString()));

                                oUpdateCotizacionesCRM.DoQuery(sGetUpdateCotizacionesCRM);

                                
                            }
                            else
                            {

                                DllFunciones.Logger(_oCompany.GetLastErrorDescription(), _PathFileLog);
                                sGetUpdateCotizacionesCRMError = sOriginalGetUpdateCotizacionesCRMError.Replace("%DocEntry%", Convert.ToString(oConsultaCotizacionesCabeceraCRM.Fields.Item("U_CP_HXC").Value.ToString())).Replace("%Estado%", "E");

                                oUpdateCotizacionesCRM.DoQuery(sGetUpdateCotizacionesCRMError);

                            }
                                                        
                            DllFunciones.liberarObjetos(oQuotation);

                            #endregion
                        }
                        else
                        {
                            DllFunciones.Logger("No existen lineas Cotizacion : " + Convert.ToString(oConsultaCotizacionesCabeceraCRM.Fields.Item("U_CP_HXC").Value.ToString()), _PathFileLog);

                            sGetUpdateCotizacionesCRMError = sOriginalGetUpdateCotizacionesCRMError.Replace("%DocEntry%", Convert.ToString(oConsultaCotizacionesCabeceraCRM.Fields.Item("U_CP_HXC").Value.ToString())).Replace("%Estado%", "E");

                            DllFunciones.Logger(sGetUpdateCotizacionesCRMError, _PathFileLog);

                            oUpdateCotizacionesCRM.DoQuery(sGetUpdateCotizacionesCRMError);
                        }

                        #endregion

                        oConsultaCotizacionesCabeceraCRM.MoveNext();

                    } while (oConsultaCotizacionesCabeceraCRM.EoF == false);

                }
                #endregion
            }
            catch (Exception e)
            {
                DllFunciones.Logger(e.ToString(), _PathFileLog);                
            }
        }
    }
}
