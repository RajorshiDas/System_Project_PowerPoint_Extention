using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Microsoft.Office.Tools;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointAddIn1
{
    public partial class ThisAddIn
    {
        private MyRibbon ribbon;
        private CustomTaskPane _qrTaskPane;

        public CustomTaskPane QRTaskPane
        {
            get { return _qrTaskPane; }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WindowSelectionChange += Application_WindowSelectionChange;

            // Create the QR Code task pane (hidden by default)
            var qrControl = new QRCodeControl();
            _qrTaskPane = this.CustomTaskPanes.Add(qrControl, "QR Code Generator");
            _qrTaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            _qrTaskPane.Width = 260;
            _qrTaskPane.Visible = false;
            _qrTaskPane.VisibleChanged += QRTaskPane_VisibleChanged;
        }

        private void QRTaskPane_VisibleChanged(object sender, EventArgs e)
        {
            // Keep the ribbon toggle button in sync when the user closes
            // the pane via the built-in X button.
            if (ribbon != null)
            {
                ribbon.SyncQRToggleButton(_qrTaskPane.Visible);
            }
        }

        private void Application_WindowSelectionChange(PowerPoint.Selection Sel)
        {
            if (ribbon != null)
            {
                ribbon.RefreshInfo();
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            this.Application.WindowSelectionChange -= Application_WindowSelectionChange;
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            ribbon = new MyRibbon();
            return Globals.Factory.GetRibbonFactory().CreateRibbonManager(
                new Microsoft.Office.Tools.Ribbon.IRibbonExtension[] { ribbon });
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
