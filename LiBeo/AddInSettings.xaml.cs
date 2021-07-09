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

            syncDBCheckBox.IsChecked = Properties.Settings.Default.SyncDBOnStartup;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.SyncDBOnStartup = (bool)syncDBCheckBox.IsChecked;
            Properties.Settings.Default.Save();
            this.Close();
        }
    }
}
