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
    /// Interaction logic for RenameFolderWindow.xaml
    /// </summary>
    public partial class RenameFolderWindow : Window
    {
        public RenameFolderWindow()
        {
            InitializeComponent();

            // display folder structure
            ThisAddIn.Structure.DisplayInTreeView(ThisAddIn.DbConn, folderExplorer, ThisAddIn.Name, false);

            // hide search suggestion list
            searchSuggestions.Visibility = Visibility.Collapsed;
            searchSuggestions.list.BorderThickness = new Thickness(0);
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                okButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            }
        }

        private void searchInput_TextChanged(object sender, TextChangedEventArgs e)
        {
            searchSuggestions.DisplaySearchSuggestions(searchInput.Text);
            folderExplorer.Visibility = searchInput.Text == "" ? Visibility.Visible : Visibility.Collapsed;
            searchSuggestions.Visibility = searchInput.Text == "" ? Visibility.Collapsed : Visibility.Visible;
        }

        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            if (newNameInput.Text == "")
                return;
            TreeViewItem selectedItem = (TreeViewItem)folderExplorer.SelectedItem;
            ListViewItem selectedSuggestedItem = (ListViewItem)searchSuggestions.SelectedItem;
            if (selectedItem == null && selectedSuggestedItem == null)
            {
                return;
            }
            int selectedId = selectedSuggestedItem == null ? (int)selectedItem.Tag : (int)selectedSuggestedItem.Tag;

            var selectedFolder = ThisAddIn.GetFolderFromPath(ThisAddIn.Structure.GetPath(ThisAddIn.DbConn, selectedId));
            if (selectedFolder == null)
            {
                MessageBox.Show("Der ausgewählte Ordner exestiert nicht mehr. Bitte synchronisieren Sie die Ordnerstruktur.",
                    "Ordner exestiert nicht mehr",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }

            try
            {
                selectedFolder.Name = newNameInput.Text;
            }
            catch(System.Runtime.InteropServices.COMException)
            {
                MessageBox.Show("Ein Ordner mit diesem Namen exestiert bereits auf derselben Ebene.",
                            "Ordner kann nicht umbenannt werden",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error);
            }

            ThisAddIn.Structure.RenameFolder(ThisAddIn.DbConn, selectedId, newNameInput.Text);
            Close();
        }

        private void cancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void searchSuggestions_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            okButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        }
    }
}
