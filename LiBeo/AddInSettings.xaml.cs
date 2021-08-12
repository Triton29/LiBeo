using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data.SQLite;

namespace LiBeo
{
    /// <summary>
    /// Interaction logic for AddInSettings.xaml
    /// </summary>
    public partial class AddInSettings : Window
    {
        public AddInSettings()
        {
            InitializeComponent();

            // display current settings
            syncDBCheckBox.IsChecked = Properties.Settings.Default.SyncFolderStructureOnStartup;
            trayPathInput.Text = Properties.Settings.Default.TrayPath;

            // add images
            trayPathButton.Content = new Image
            {
                Source = new BitmapImage(new Uri(AppDomain.CurrentDomain.BaseDirectory + @"\img\folder.png"))
            };
        }

        /// <summary>
        /// Called when ok button is pressed; saves all settings
        /// </summary>
        private void OKButton_Click(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.SyncFolderStructureOnStartup = (bool)syncDBCheckBox.IsChecked;
            Properties.Settings.Default.TrayPath = trayPathInput.Text;
            Properties.Settings.Default.Save();
            this.Close();
        }

        /// <summary>
        /// Called when cancel button is pressed; closes the window
        /// </summary>
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Called when the button next to the tray path input is pressed; opens a window to select a folder
        /// </summary>
        private void trayPathButton_Click(object sender, RoutedEventArgs e)
        {
            SelectFolder selectFolderWindow = new SelectFolder();
            if(selectFolderWindow.ShowDialog() == false && !selectFolderWindow.Canceled)
            {
                string trayPath = "";
                foreach(string folder in selectFolderWindow.SelectedFolderPath)
                {
                    trayPath = trayPath + @"\" + folder;
                }
                trayPathInput.Text = trayPath;
            }
        }

        /// <summary>
        /// Called when the quickAccessListButton is pressed; 
        /// opens a window to select multiple folders for the quick access list
        /// </summary>
        private void quickAccessListButton_Click(object sender, RoutedEventArgs e)
        {
            ThisAddIn.DbConn.Open();
            SQLiteCommand DbCmd = new SQLiteCommand(ThisAddIn.DbConn);

            DbCmd.CommandText = "SELECT * FROM quick_access_folders";
            SQLiteDataReader dataReader = DbCmd.ExecuteReader();
            List<int> ids = new List<int>();
            while (dataReader.Read())
            {
                ids.Add(dataReader.GetInt32(0));
            }
            dataReader.Close();

            ThisAddIn.DbConn.Close();

            MultiSelectFolder multiSelectFolderWindow = new MultiSelectFolder();
            multiSelectFolderWindow.PreSelectedFolderIds = ids;

            if (multiSelectFolderWindow.ShowDialog() == false && !multiSelectFolderWindow.Canceled)
            {
                ThisAddIn.DbConn.Open();
                DbCmd.CommandText = "DELETE FROM quick_access_folders";
                DbCmd.ExecuteNonQuery();

                foreach(int id in multiSelectFolderWindow.SelectedFolderIds)
                {
                    DbCmd.CommandText = "INSERT OR IGNORE INTO quick_access_folders (folder) VALUES (@id)";
                    DbCmd.Parameters.AddWithValue("@id", id);
                    DbCmd.Prepare();
                    DbCmd.ExecuteNonQuery();
                }
                ThisAddIn.DbConn.Close();
            }
        }
    }
}
