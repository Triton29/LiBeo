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
using System.Windows.Forms;
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
            if (ThisAddIn.GetSetting<int>("sync_db") == 1)
                syncDBCheckBox.IsChecked = true;
            dbInput.Text = Properties.Settings.Default.DbPath;
            stopWordsInput.Text = ThisAddIn.GetSetting<string>("stop_words_path");
            trayPathInput.Text = ThisAddIn.GetSetting<string>("tray_path");

            // add images
            dbButton.Content = new Image
            {
                Source = new BitmapImage(new Uri(AppDomain.CurrentDomain.BaseDirectory + @"\img\folder.png"))
            };
            stopWordsButton.Content = new Image
            {
                Source = new BitmapImage(new Uri(AppDomain.CurrentDomain.BaseDirectory + @"\img\folder.png"))
            };
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
            if(syncDBCheckBox.IsChecked == true)
                ThisAddIn.SetSetting<int>("sync_db", 1);
            else
                ThisAddIn.SetSetting<int>("sync_db", 0);
            Properties.Settings.Default.DbPath = dbInput.Text;
            Properties.Settings.Default.Save();
            ThisAddIn.SetSetting<string>("stop_words_path", stopWordsInput.Text);
            ThisAddIn.SetSetting<string>("tray_path", trayPathInput.Text);
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
            SQLiteCommand DbCmd = new SQLiteCommand(ThisAddIn.DbConn);

            DbCmd.CommandText = "SELECT * FROM quick_access_folders";
            SQLiteDataReader dataReader = DbCmd.ExecuteReader();
            List<int> ids = new List<int>();
            while (dataReader.Read())
            {
                ids.Add(dataReader.GetInt32(0));
            }
            dataReader.Close();

            MultiSelectFolder multiSelectFolderWindow = new MultiSelectFolder();
            multiSelectFolderWindow.PreSelectedFolderIds = ids;

            if (multiSelectFolderWindow.ShowDialog() == false && !multiSelectFolderWindow.Canceled)
            {
                DbCmd.CommandText = "DELETE FROM quick_access_folders";
                DbCmd.ExecuteNonQuery();

                foreach(int id in multiSelectFolderWindow.SelectedFolderIds)
                {
                    DbCmd.CommandText = "INSERT OR IGNORE INTO quick_access_folders (folder) VALUES (@id)";
                    DbCmd.Parameters.AddWithValue("@id", id);
                    DbCmd.Prepare();
                    DbCmd.ExecuteNonQuery();
                }
            }
        }

        private void stopWordsButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            DialogResult res = fd.ShowDialog();
            fd.Filter = "Text|*.txt";
            if (!string.IsNullOrWhiteSpace(fd.FileName))
            {
                stopWordsInput.Text = (fd.FileName);
            }
        }

        private void dbButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            fd.Filter = "DB-Datei|*.db";
            DialogResult res = fd.ShowDialog();
            if (!string.IsNullOrWhiteSpace(fd.FileName))
            {
                dbInput.Text = (fd.FileName);
            }
        }
    }
}
