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
                Properties.Settings.Default.TrayPath = trayPath;
                Properties.Settings.Default.Save();
            }
        }
    }
}
