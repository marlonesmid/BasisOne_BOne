using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Reflection;
using System.Xml;
using System.IO;
using Funciones;
//using BOCore;

namespace BOBillingConsolidator
{
    public class BillingConsolidator
    {

        public void AddMenu_BillingConsolidator(SAPbouiCOM.MenuCreationParams _oCreationPackage, SAPbouiCOM.Menus _oMenus, SAPbouiCOM.MenuItem _oMenuItem, SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _company , string AddIn, string _sNameDB )
        {

            try
            {
                #region Variables y Objetos

                string sQuerieValidacion = null;
                string LicenciaAddIn = null;
                string ParametrosLicencia = null;
                string LicenciaValida = null;
                string sPath = null;
                string sMotor = null;

                XmlDocument XmlConsulta = XmlConsulta = new XmlDocument();

                sPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                sMotor = Convert.ToString(_company.DbServerType);

                SAPbobsCOM.Recordset oValidacion = ((SAPbobsCOM.Recordset)_company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));

                #endregion

                #region Instanciacion Dll

                Funciones.Comunes DllFunciones = new Funciones.Comunes();
               // BOCore.Core DllCore = new BOCore.Core();

                #endregion

                #region Menu Consolidacion Facturación

                _oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                _oCreationPackage.UniqueID = "mnuHX_BC";
                _oCreationPackage.Enabled = true;
                _oCreationPackage.String = "Consolidador Facturación";
                _oCreationPackage.Image = "";
                _oCreationPackage.Position = _oCreationPackage.Position + 1;
                _oMenus = _oMenuItem.SubMenus;
                _oMenus.AddEx(_oCreationPackage);           

                #endregion
 
            }

            catch (Exception)
            {

                throw;
            }
        }

        public string VersionAddOn()
        {
            try
            {

                Assembly Assembly = Assembly.LoadFrom("BOBillingConsolidator.dll");
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
