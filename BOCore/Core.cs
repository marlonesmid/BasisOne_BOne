using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Funciones;
using SAPbobsCOM;
using SAPbouiCOM;
using System.IO;

namespace BOCore
{
    public class Core
    {
        #region Instanciacion Dll's

        Funciones.Comunes DllFunciones = new Funciones.Comunes();
        
        #endregion

        public void LoadParametersFormGestorAddOn(SAPbouiCOM.Application _sboapp, SAPbobsCOM.Company _oCompany, SAPbouiCOM.Form _oFormGA, string sVersionInstaladorAddOn)
        {
            try
            {

                #region Creacion Variables y Objetos

                string sGridAddInAvailable = null;
                string sGridAddInActive = null;

                SAPbouiCOM.DataTable oDtAddInAvailable;
                SAPbouiCOM.DataTable oDtAddInActive;

                SAPbouiCOM.Grid oGridAddInAvailable;
                SAPbouiCOM.Grid oGridAddInActive;

                SAPbouiCOM.PictureBox oLogoBO;

                SAPbouiCOM.Folder oFolder1;

                SAPbouiCOM.StaticText olblVersion;

                #endregion

                #region Instanciacion de variables y objetos

                oLogoBO = (SAPbouiCOM.PictureBox)(_oFormGA.Items.Item("oLogo").Specific);
                oDtAddInAvailable = _oFormGA.DataSources.DataTables.Add("oDtGridAD");
                oDtAddInActive = _oFormGA.DataSources.DataTables.Add("oDtGridAA");

                oGridAddInAvailable = (SAPbouiCOM.Grid)(_oFormGA.Items.Item("GridDispo").Specific);
                oGridAddInActive = (SAPbouiCOM.Grid)(_oFormGA.Items.Item("GridActi").Specific);

                oFolder1 = (SAPbouiCOM.Folder)_oFormGA.Items.Item("Folder1").Specific;

                olblVersion = (SAPbouiCOM.StaticText)(_oFormGA.Items.Item("lblVersion").Specific);

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

                #region Asignacion Logo

                oLogoBO.Picture = (Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Core\\Images\\LogoBO20x20.bmp");

                #endregion

                #region Asignacion Label

                olblVersion.Caption = "AddOn BOne " + sVersionInstaladorAddOn;
                olblVersion.Item.TextStyle = 1;

                #endregion

                _oFormGA.Visible = true;

                oFolder1.Select();

                _oFormGA.Refresh();

            }
            catch (Exception e)
            {

                DllFunciones.sendErrorMessage(_sboapp, e);
            }





        }
    }
}
