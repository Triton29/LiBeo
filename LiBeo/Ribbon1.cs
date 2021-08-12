using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Data.SQLite;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace LiBeo
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon1()
        {
            
        }

        /// <summary>
        /// Creates a window with the sort-actions for the selected mails
        /// </summary>
        /// <param name="control">The button which calls this function</param>
        public void AutoSort_Click(Office.IRibbonControl control)
        {
            Actions actionsForm = new Actions();
            actionsForm.Show();
            // select auto sort tab
            actionsForm.tabConrol.SelectedIndex = 0;
        }

        /// <summary>
        /// Creates a window with the sort-actions for the selected mails
        /// </summary>
        /// <param name="control">The button which calls this function</param>
        public void ManualSort_Click(Office.IRibbonControl control)
        {
            Actions actionsForm = new Actions();
            actionsForm.Show();
            // select manual sort tab
            actionsForm.tabConrol.SelectedIndex = 1;
        }

        /// <summary>
        /// Creates a window with the sort-actions for the selected mails
        /// </summary>
        /// <param name="control">The button which calls this function</param>
        public void CreateDir_Click(Office.IRibbonControl control)
        {
            Actions actionsForm = new Actions();
            actionsForm.Show();
            // select create directory tab
            actionsForm.tabConrol.SelectedIndex = 2;
            actionsForm.quickAccessList.TabIndex = 0;
        }

        /// <summary>
        /// Moves the selected mails to the tray
        /// </summary>
        /// <param name="control">The button which calls this function</param>
        public void MoveToTray_Click(Office.IRibbonControl control)
        {
            Actions.MoveToTray();
        }

        /// <summary>
        /// Synchronizes the database where the folder structure is saved with the current folder structure
        /// </summary>
        /// <param name="control">The button which calls this function</param>
        public void SyncFolderStructure(Office.IRibbonControl control)
        {
            WaitWindow waitWindow = new WaitWindow();
            waitWindow.Show();

            ThisAddIn.SyncFolderStructure();

            waitWindow.Close();
        }

        /// <summary>
        /// Synchronizes the database with the stop words list (stop_words.txt)
        /// </summary>
        /// <param name="control"></param>
        public void SyncStopWords(Office.IRibbonControl control)
        {
            WaitWindow waitWindow = new WaitWindow();
            waitWindow.Show();

            ThisAddIn.SyncStopWords();

            waitWindow.Close();
        }

        /// <summary>
        /// Creates a window with settings for this Add-In
        /// </summary>
        /// <param name="control">The button which calls this function</param>
        public void AddInSettings(Office.IRibbonControl control)
        {
            AddInSettings settingsWindow = new AddInSettings();
            settingsWindow.Show();
        }

        /// <summary>
        /// Creates a window with info for this Add-In
        /// </summary>
        /// <param name="control">The button which calls this function</param>
        public void AddInInfo(Office.IRibbonControl control)
        {
            AddInInfo infoForm = new AddInInfo();
            infoForm.Show();
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("LiBeo.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
