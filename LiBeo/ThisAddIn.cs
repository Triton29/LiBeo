// LiBeo @ 2021 Leo Mühlböck
// LiBeo = Litteras diribeo (latin) = mail sort

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Data.SQLite;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace LiBeo
{
    public partial class ThisAddIn
    {
        public static string Version = "0.1 (Alpha)";
        public static string DbName = AppDomain.CurrentDomain.BaseDirectory + @"\data.db";

        public static Outlook.Folder RootFolder { get; set; }
        public static FolderStructure Structure { get; set; }

        /// <summary>
        /// Create the Ribbon
        /// </summary>
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // initialize properties
            RootFolder = (Outlook.Folder) this.Application.ActiveExplorer().Session.DefaultStore.GetRootFolder();
            Structure = new FolderStructure(RootFolder);

            // syncronize database if enabled
            if(Properties.Settings.Default.SyncDBOnStartup)
                SyncDatabase();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            
        }

        /// <summary>
        /// Synchronizes the database where the folder structure is saved with the current folder structure
        /// </summary>
        public static void SyncDatabase()
        {
            SQLiteConnection dbConn = new SQLiteConnection("Data Source=" + DbName);
            dbConn.Open();
            Structure.SaveToDB(dbConn);
            dbConn.Close();
        }

        /// <summary>
        /// Gets all selected mails and returns them
        /// </summary>
        /// <returns>All selected mails</returns>
        internal static IEnumerable<Outlook.MailItem> GetSelectedMails()
        {
            foreach(Outlook.MailItem mail in new Outlook.Application().ActiveExplorer().Selection)
            {
                yield return mail;
            }
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
