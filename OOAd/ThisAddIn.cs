using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OOAd
{
    public partial class ThisAddIn
    {
        Office.CommandBar newToolBar;
        Office.CommandBarButton firstButton;
        Office.CommandBarButton secondButton;
        Outlook.Explorers selectExplorers;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            selectExplorers = this.Application.Explorers;
            selectExplorers.NewExplorer += new Outlook
                .ExplorersEvents_NewExplorerEventHandler(newExplorer_Event);
            AddToolbar();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
        }
        private void newExplorer_Event(Outlook.Explorer new_Explorer)
        {
            ((Outlook._Explorer)new_Explorer).Activate();
            //Office.CommandBars cmdBars =
            //        this.Application.ActiveExplorer().CommandBars;
            
            newToolBar = null;//.Add("NewToolBar",Office.MsoBarPosition.msoBarTop, false, true); ;
            AddToolbar();
        }

        private void AddToolbar()
        {

            if (newToolBar == null)
            {
                Office.CommandBars cmdBars =
                    this.Application.ActiveExplorer().CommandBars;
                foreach (Outlook.Explorer cmdbar in this.Application.ActiveExplorer().Application.Explorers)
                {
                    Debug.WriteLine(cmdbar.CommandBars);
                }
                newToolBar = cmdBars["Meeting Workspace"];//.Add("NewToolBar",
                   // Office.MsoBarPosition.msoBaкTop, false, true);
            }
            try
            {

                Office.CommandBarButton button_1 =
                    (Office.CommandBarButton)newToolBar.Controls
                    .Add(1, missing, missing, missing, missing);
                button_1.Style = Office
                    .MsoButtonStyle.msoButtonCaption;
                button_1.Caption = "Button 1";
                button_1.Tag = "Button1";
                if (this.firstButton == null)
                {
                    this.firstButton = button_1;
                    firstButton.Click += new Office.
                        _CommandBarButtonEvents_ClickEventHandler
                        (ButtonClick);
                }

                Office.CommandBarButton button_2 = (Office
                    .CommandBarButton)newToolBar.Controls.Add
                    (1, missing, missing, missing, missing);
                button_2.Style = Office
                    .MsoButtonStyle.msoButtonCaption;
                button_2.Caption = "Button 2";
                button_2.Tag = "Button2";
                newToolBar.Visible = true;
                if (this.secondButton == null)
                {
                    this.secondButton = button_2;
                    secondButton.Click += new Office.
                        _CommandBarButtonEvents_ClickEventHandler
                        (ButtonClick);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonClick(Office.CommandBarButton ctrl,
                ref bool cancel)
        {
            MessageBox.Show("You clicked: " + ctrl.Caption);
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
