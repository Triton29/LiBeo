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
    /// Interaction logic for SelectFolder.xaml
    /// </summary>
    public partial class SelectFolder : Window
    {
        public List<string> SelectedFolderPath;
        public bool Canceled = false;

        public SelectFolder()
        {
            InitializeComponent();

            // display folder structure in tree view
            ThisAddIn.DbConn.Open();
            ThisAddIn.Structure.DisplayInTreeView(ThisAddIn.DbConn, folderExplorer);
            ThisAddIn.DbConn.Close();
        }

        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Called when any key on the keyboard is pressed; implements shortcuts
        /// </summary>
        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                okButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            }
        }

        /// <summary>
        /// Called when cancel button is pressed; closes the window
        /// </summary>
        private void cancelButton_Click(object sender, RoutedEventArgs e)
        {
            Canceled = true;
            this.Close();
        }

        /// <summary>
        /// Called when the selected item in the folder explorer treeview has changed; 
        /// writes the currently selected folder in the public variable
        /// </summary>
        private void folderExplorer_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            ThisAddIn.DbConn.Open();
            int id = (int)((TreeViewItem)folderExplorer.SelectedItem).Tag;
            SelectedFolderPath = FolderStructure.GetPath(ThisAddIn.DbConn, id);
            ThisAddIn.DbConn.Close();
        }
    }
}
