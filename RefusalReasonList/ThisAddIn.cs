using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Reflection;
using System.Windows.Forms;
using DocumentFormat.OpenXml.ExtendedProperties;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using DocumentFormat.OpenXml.Office2013.WebExtentionPane;
using System.Diagnostics;
using OpenXmlPowerTools;

namespace RefusalReasonList
{
    public partial class ThisAddIn
    {
        public Microsoft.Office.Tools.CustomTaskPane myPane;

        private UserControlAccount userControl1;
        private Microsoft.Office.Tools.CustomTaskPane taskPaneValue;
        private Dictionary<string, Microsoft.Office.Tools.CustomTaskPaneCollection> panesDictionary = new Dictionary<string, Microsoft.Office.Tools.CustomTaskPaneCollection>();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            myPane = this.CustomTaskPanes.Add(new UserControlAccount(), "APIアカウント");
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void taskPaneValue_VisibleChanged(object sender, System.EventArgs e)
        {
            if (this.panesDictionary.Keys.Count == 0)
            {
            }
            else // Need to do more check if there are more than one key
            {
            }
        }
        public Microsoft.Office.Tools.CustomTaskPane TaskPane
        {
            get
            {
                Microsoft.Office.Tools.CustomTaskPaneCollection panes;

                if (this.panesDictionary.ContainsKey(this.FileName))
                {
                    panes = this.panesDictionary[this.FileName];
                    taskPaneValue = panes[0];  // we only added one for each file
                }
                else
                {
                    if (this.panesDictionary.Count == 0)
                    {
                        panes = this.CustomTaskPanes;
                    }
                    else
                    {
                        panes = Globals.Factory.CreateCustomTaskPaneCollection(null, null, this.FileName, "CustomTaskPanes", this);
                    }
                    panesDictionary.Add(this.FileName, panes);
                    userControl1 = new UserControlAccount();
                    //taskPaneValue = panes.Add(userControl1, this.FileName + " - MyCustomTaskPane");
                    taskPaneValue = panes.Add(userControl1, "APIアカウント");
                    taskPaneValue.VisibleChanged += new EventHandler(taskPaneValue_VisibleChanged);
                }
                return taskPaneValue;
            }
        }
        public string FileName
        {
            get
            {
                if (Application.ActiveWorkbook == null)
                {
                    return "defaultcustomtaskpane";
                }
                else
                {
                    return Application.ActiveWorkbook.FullNameURLEncoded;
                }
            }
        }
        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
