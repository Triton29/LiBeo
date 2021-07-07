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
        public static string version = "0.1 (Alpha)";

        public static Outlook.Folder rootFolder { get; set; }
        public static string dbName = @"\data.db";

        /// <summary>
        /// Create the Ribbon
        /// </summary>
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            rootFolder = (Outlook.Folder) this.Application.ActiveExplorer().Session.DefaultStore.GetRootFolder();

            SQLiteConnection dbConn = new SQLiteConnection("Data Source=" + AppDomain.CurrentDomain.BaseDirectory + dbName);
            dbConn.Open();
            FolderStructure folderStructure = new FolderStructure(rootFolder);
            folderStructure.SaveToDB(dbConn);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            
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
