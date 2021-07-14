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

            syncDBCheckBox.IsChecked = Properties.Settings.Default.SyncFolderStructureOnStartup;
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
    }
}
