﻿// LiBeo @ 2021 Leo Mühlböck
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
        public static string Version = "1.0";
        public static string DbPath = Properties.Settings.Default.DbPath;
        public static string StopWordsPath { get; set; }
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
                SyncFolderStructure();
                //SyncStopWords();
            }
        }

        /// <summary>
        /// Runs action for the WaitThread; creates a wait window
        /// </summary>
        static void WaitThreadAction()
        {
            WaitWindow waitWindow = new WaitWindow();
            waitWindow.ShowDialog();
        }
        /// <summary>
        /// Shows a wait window until CloseWaitWindow() method is called
        /// </summary>
        /// <returns>The wait thread (neccessary for closing the window)</returns>
        public static Thread ShowWaitWindow()
        {
            ThreadStart threadStart = new ThreadStart(WaitThreadAction);
            Thread waitThread = new Thread(threadStart);
            waitThread.SetApartmentState(ApartmentState.STA);
            waitThread.Start();
            return waitThread;
        }
        /// <summary>
        /// Closes the wait window created in a wait thread
        /// </summary>
        /// <param name="waitThread">The wait thread returned by ShowWaitWindow() method</param>
        public static void CloseWaitWindow(Thread waitThread)
        {
            try
            {
                waitThread.Abort();
            }
            catch { }
        }
        #endregion

        /// <summary>
        /// Called when the Add-In starts up; sets up all properties and the database
        /// </summary>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // initialize properties
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

            StopWordsPath = GetSetting<string>("stop_words_path");
            if (!StopWordsPath.Contains("\\"))
                StopWordsPath = AppDomain.CurrentDomain.BaseDirectory + StopWordsPath;

            // sync folder structure and stop words in new thread because it takes a long time
            ThreadStart threadStart = new ThreadStart(StartupThreadAction);
            Thread startupThread = new Thread(threadStart);
            startupThread.Start();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            DbConn.Close();
        }

        /// <summary>
        /// Synchronizes the database where the folder structure is saved with the current folder structure
        /// </summary>
        public static void SyncFolderStructure()
        {
            Structure.SaveToDB(DbConn);
        }

        /// <summary>
        /// Synchronizes the database with the stop words list (stop_words.txt)
        /// </summary>
        public static void SyncStopWords()
        {
            try
            {
                SQLiteCommand dbCmd = new SQLiteCommand(DbConn);

                System.IO.StreamReader file = new System.IO.StreamReader(StopWordsPath);
                string line;
                while ((line = file.ReadLine()) != null)
                {
                    dbCmd.CommandText = "INSERT OR IGNORE INTO stop_words VALUES (@word)";
                    dbCmd.Parameters.AddWithValue("@word", line);
                    dbCmd.Prepare();
                    dbCmd.ExecuteNonQuery();
                }
            }
            catch
            {
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
                    folder = (Outlook.Folder) folder.Folders[f];
            }
            return folder;
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
                    return (T)Convert.ChangeType(dataReader.GetInt32(0), typeof(T));
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
