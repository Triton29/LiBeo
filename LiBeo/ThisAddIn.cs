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
using System.Threading;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace LiBeo
{
    public partial class ThisAddIn
    {
        public static string Version = "1.3";
        public static string DbPathTxt = AppDomain.CurrentDomain.BaseDirectory + "db_path.txt";
        public static string DbPath { get; set; }
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

        #region thread methods
        /// <summary>
        /// Runs action for the StartupThread; synchronises the database while outlook has already started
        /// </summary>
        void StartupThreadAction()
        {
            Thread.Sleep(new TimeSpan(0, 0, 1));

            // synchronize folder structure if enabled
            if (GetSetting<int>("sync_db") == 1)
            {
                SyncFolderStructure(false);
                SyncStopWords(false);
            }
        }

        /// <summary>
        /// Creates a wait window running in another thread
        /// </summary>
        /// <returns>The created wait window</returns>
        public static WaitWindow CreateWaitWindow()
        {
            WaitWindow waitWindow = null;
            Thread waitThread = new Thread(() => 
            {
                waitWindow = new WaitWindow();
                waitWindow.Show();
                System.Windows.Threading.Dispatcher.Run();
            });
            waitThread.SetApartmentState(ApartmentState.STA);
            waitThread.IsBackground = true;
            waitThread.Start();
            while (waitWindow == null)
                Thread.Sleep(10);
            return waitWindow;
        }
        /// <summary>
        /// Closes a wait window running in another thread
        /// </summary>
        /// <param name="waitWindow">The wait window that should be closed</param>
        /// <returns>If the wait window could be closed</returns>
        public static bool CloseWaitWindow(WaitWindow waitWindow)
        {
            for (int i = 0; waitWindow == null; i++)
            {
                Thread.Sleep(10);
                if (i > 1000)
                    return false;
            }
            waitWindow.Dispatcher.Invoke(() => { waitWindow.Close(); });
            return true;
        }
        #endregion

        /// <summary>
        /// Called when the Add-In starts up; sets up all properties and the database
        /// </summary>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                // initialize properties
                DbPath = GetDbPath();
                RootFolder = (Outlook.Folder)Application.ActiveExplorer().Session.DefaultStore.GetRootFolder();
                EmailAddress = Application.ActiveExplorer().Session.CurrentUser.Address;
                Name = Application.Session.Accounts[1].DisplayName;
                Structure = new FolderStructure(RootFolder);

                if (!DbPath.Contains("\\"))
                    DbPath = AppDomain.CurrentDomain.BaseDirectory + DbPath;
                DbConn = new SQLiteConnection("Data Source=" + DbPath);
                DbConn.Open();

                // setup database
                SetupDatabase();

                // sync folder structure and stop words in new thread because it takes a long time
                ThreadStart threadStart = new ThreadStart(StartupThreadAction);
                Thread startupThread = new Thread(threadStart);
                startupThread.Start();
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                MessageBox.Show("Bitte überprüfen Sie die Netzwerkverbindung",
                    "LiBeo konnte nicht geladen werden",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Called when Add-In shuts down; closes db connection
        /// </summary>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            DbConn.Close();
        }

        /// <summary>
        /// Gets the database path out of the DbPathTxt-file
        /// </summary>
        /// <returns>The database path</returns>
        public static String GetDbPath()
        {
            try
            {
                return System.IO.File.ReadAllText(DbPathTxt);
            }
            catch (System.IO.FileNotFoundException)
            {
                SetDbPath("data.db");
                return "data.db";
            }
        }

        /// <summary>
        /// Sets the database path to the DbPathTxt-file
        /// </summary>
        /// <param name="path">The new database path</param>
        public static void SetDbPath (string path)
        {
            System.IO.File.WriteAllText(DbPathTxt, path);
        }

        /// <summary>
        /// Synchronizes the database where the folder structure is saved with the current folder structure
        /// </summary>
        public static void SyncFolderStructure(bool createWaitWindow=true)
        {
            WaitWindow waitWindow = null;
            if (createWaitWindow)
                waitWindow = CreateWaitWindow();
            Structure.SaveToDB(DbConn);
            if (createWaitWindow)
                CloseWaitWindow(waitWindow);
        }

        /// <summary>
        /// Synchronizes the database with the stop words list (stop_words.txt)
        /// <param name="createWaitWindow">If a wait window should be created</param>
        /// </summary>
        public static void SyncStopWords(bool createWaitWindow=true)
        {
            WaitWindow waitWindow = null;
            if(createWaitWindow)
                waitWindow = CreateWaitWindow();
            try
            {
                string path = GetSetting<string>("stop_words_path");
                if (!path.Contains("\\"))
                    path = AppDomain.CurrentDomain.BaseDirectory + path;
                SQLiteCommand dbCmd = new SQLiteCommand(DbConn);

                System.IO.StreamReader file = new System.IO.StreamReader(path);
                string line;
                while ((line = file.ReadLine()) != null)
                {
                    dbCmd.CommandText = "INSERT OR IGNORE INTO stop_words VALUES (@word)";
                    dbCmd.Parameters.AddWithValue("@word", line);
                    dbCmd.Prepare();
                    dbCmd.ExecuteNonQuery();
                }

                if(createWaitWindow)
                    CloseWaitWindow(waitWindow);
            }
            catch (System.IO.FileNotFoundException)
            {
                if (createWaitWindow)
                    CloseWaitWindow(waitWindow);
                MessageBox.Show("Beim Synchronisieren der Stop Words ist etwas schiefgelaufen. Bitte überprüfen Sie den Pfad.",
                            "Stop Words können nicht synchronisiert werden",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error);
            }
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
                {
                    try
                    {
                        folder = (Outlook.Folder)folder.Folders[f];
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        return null;
                    }
                }
            }
            return folder;
        }

        public static Outlook.Folder GetFolderFromPath(List<string> path)
        {
            return GetFolderFromPath(string.Join("\\", path));
        }

        /// <summary>
        /// Sets up the database for quick access list, folder structure, stop words and folder
        /// </summary>
        static void SetupDatabase()
        {
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

            dbCmd.CommandText = "CREATE TABLE IF NOT EXISTS settings (name varchar(255) UNIQUE, value_int int, value_str varchar(255))";
            dbCmd.ExecuteNonQuery();
            // settings
            dbCmd.CommandText = "INSERT OR IGNORE INTO settings (name, value_int) VALUES ('sync_db', 1)";
            dbCmd.ExecuteNonQuery();
            dbCmd.CommandText = "INSERT OR IGNORE INTO settings (name, value_int, value_str) VALUES ('stop_words_path', -1, 'stop_words.txt')";
            dbCmd.ExecuteNonQuery();
            dbCmd.CommandText = "INSERT OR IGNORE INTO settings (name, value_int, value_str) VALUES ('tray_path', -1, '\\Ablage')";
            dbCmd.ExecuteNonQuery();
            dbCmd.CommandText = "INSERT OR IGNORE INTO settings (name, value_int) VALUES ('history_limit', 10)";
            dbCmd.ExecuteNonQuery();
        }

        /// <summary>
        /// Gets a setting saved in the database
        /// </summary>
        /// <typeparam name="T">The datatype of the setting (int or string)</typeparam>
        /// <param name="settingName">The name of the setting</param>
        /// <returns>The value of the setting as given datatype</returns>
        public static T GetSetting<T>(string settingName)
        {
            SQLiteCommand dbCmd = new SQLiteCommand(DbConn);
            dbCmd.CommandText = "SELECT value_int, value_str FROM settings WHERE name=@name";
            dbCmd.Parameters.AddWithValue("name", settingName);
            dbCmd.Prepare();
            SQLiteDataReader dataReader = dbCmd.ExecuteReader();
            if (dataReader.Read())
            {
                if(dataReader.GetInt32(0) == -1)
                {
                    return (T) Convert.ChangeType(dataReader.GetString(1), typeof(T));
                }
                else
                {
                    return (T) Convert.ChangeType(dataReader.GetInt32(0), typeof(T));
                }
            }
            return default(T);
        }
        /// <summary>
        /// Sets a setting saved in the database to a new value
        /// </summary>
        /// <typeparam name="T">The datatype of the setting (int or string)</typeparam>
        /// <param name="settingName">The name of the setting</param>
        /// <param name="value">The new value as given datatype</param>
        public static void SetSetting<T>(string settingName, T value)
        {
            SQLiteCommand dbCmd = new SQLiteCommand(DbConn);
            if(value is int)
            {
                dbCmd.CommandText = "UPDATE settings SET value_int=@value WHERE name=@name";
                dbCmd.Parameters.AddWithValue("@value", value);
                dbCmd.Parameters.AddWithValue("@name", settingName);
                dbCmd.Prepare();
                dbCmd.ExecuteNonQuery();
            }
            if(value is string)
            {
                dbCmd.CommandText = "UPDATE settings SET value_str=@value WHERE name=@name";
                dbCmd.Parameters.AddWithValue("@value", value);
                dbCmd.Parameters.AddWithValue("@name", settingName);
                dbCmd.Prepare();
                dbCmd.ExecuteNonQuery();
            }
        }

        #region enumerable methods
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
        #endregion

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
