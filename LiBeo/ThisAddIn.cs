// LiBeo @ 2021 Leo Mühlböck
// LiBeo = Litteras diribeo (latin) = mail sort

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Windows;
using System.Windows.Media;
using System.Data.SQLite;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace LiBeo
{
    public partial class ThisAddIn
    {
        public static string Version = "1.0";
        public static string DbPath = AppDomain.CurrentDomain.BaseDirectory + @"\data.db";
        public static string StopWordsPath = AppDomain.CurrentDomain.BaseDirectory + @"\stop_words.txt";

        public static Outlook.Folder RootFolder { get; set; }
        public static string EmailAddress { get; set; }
        public static string Name { get; set; }
        public static FolderStructure Structure { get; set; }
        public static SQLiteConnection DbConn { get; set; }

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
            EmailAddress = this.Application.ActiveExplorer().Session.CurrentUser.Address;
            Name = this.Application.Session.Accounts[1].DisplayName;
            Structure = new FolderStructure(RootFolder);
            DbConn = new SQLiteConnection("Data Source=" + DbPath);

            // synchronize folder structure if enabled
            if (Properties.Settings.Default.SyncFolderStructureOnStartup)
                SyncFolderStructure();

            // setup database
            SetupDatabase();

            // synchronizes stop words if not done yet
            if (!Properties.Settings.Default.SyncedStopWords)
            {
                SyncStopWords();
                Properties.Settings.Default.SyncedStopWords = true;
                Properties.Settings.Default.Save();
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            
        }

        /// <summary>
        /// Synchronizes the database where the folder structure is saved with the current folder structure
        /// </summary>
        public static void SyncFolderStructure()
        {
            DbConn.Open();
            Structure.SaveToDB(DbConn);
            DbConn.Close();
        }

        /// <summary>
        /// Synchronizes the database with the stop words list (stop_words.txt)
        /// </summary>
        public static void SyncStopWords()
        {
            DbConn.Open();
            SQLiteCommand dbCmd = new SQLiteCommand(DbConn);

            System.IO.StreamReader file = new System.IO.StreamReader(StopWordsPath);
            string line;
            while((line = file.ReadLine()) != null){
                dbCmd.CommandText = "INSERT OR IGNORE INTO stop_words VALUES (@word)";
                dbCmd.Parameters.AddWithValue("@word", line);
                dbCmd.Prepare();
                dbCmd.ExecuteNonQuery();
            }

            DbConn.Close();
        }

        /// <summary>
        /// Gets and returns a folder based on a path
        /// </summary>
        /// <param name="path">The path of the folder</param>
        /// <returns>The folder based on the path</returns>
        public static Outlook.Folder GetFolderFromPath(string path)
        {
            Outlook.Folder folder = RootFolder;
            foreach(string f in path.Split('\\'))
            {
                if(f != "")
                    folder = (Outlook.Folder) folder.Folders[f];
            }
            return folder;
        }

        /// <summary>
        /// Sets up the database for quick access list, folder structure, stop words and folder
        /// </summary>
        public static void SetupDatabase()
        {
            DbConn.Open();
            SQLiteCommand dbCmd = new SQLiteCommand(DbConn);

            dbCmd.CommandText = "CREATE TABLE IF NOT EXISTS quick_access_folders (folder int UNIQUE)";
            dbCmd.ExecuteNonQuery();

            dbCmd.CommandText =
                "CREATE TABLE IF NOT EXISTS folders (" +
                "name varchar(255), id INTEGER PRIMARY KEY AUTOINCREMENT, parent_id int, got_deleted bit, UNIQUE(name, parent_id))";
            dbCmd.ExecuteNonQuery();

            dbCmd.CommandText = "CREATE TABLE IF NOT EXISTS stop_words (word varchar(255) UNIQUE)";
            dbCmd.ExecuteNonQuery();

            dbCmd.CommandText = "CREATE TABLE IF NOT EXISTS tags (folder int, tag varchar(255), UNIQUE(folder, tag))";
            dbCmd.ExecuteNonQuery();
            dbCmd.CommandText = "CREATE TABLE IF NOT EXISTS current_mail_subject (folder int, word varchar(255) UNIQUE)";
            dbCmd.ExecuteNonQuery();

            DbConn.Close();
        }

        /// <summary>
        /// Gets all selected mails and returns them
        /// </summary>
        /// <returns>All selected mails</returns>
        internal static IEnumerable<Outlook.MailItem> GetSelectedMails()
        {
            foreach (object mail in new Outlook.Application().ActiveExplorer().Selection)
            {
                if(mail is Outlook.MailItem)
                    yield return (Outlook.MailItem) mail;
            }
        }

        /// <summary>
        /// Gets all logical children of a wpf element
        /// </summary>
        /// <returns>All children of a wpf element</returns>
        internal static IEnumerable<T> GetLogicalChildren<T>(DependencyObject depObj) where T : DependencyObject
        {
            if (depObj == null)
                yield return null;
            
            foreach (var rawChild in LogicalTreeHelper.GetChildren(depObj))
            {
                if(rawChild is DependencyObject)
                {
                    var child = (DependencyObject)rawChild;
                    if (child is T)
                        yield return (T)child;
                    foreach (T childOfChild in GetLogicalChildren<T>(child))
                        yield return childOfChild;
                }
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
