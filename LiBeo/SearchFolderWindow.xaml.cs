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
    /// Interaction logic for SearchFolderWindow.xaml
    /// </summary>
    public partial class SearchFolderWindow : Window
    {
        public SearchFolderWindow()
        {
            InitializeComponent();

            suggestionList.DisplayListViewMsg("Geben Sie einen Suchbegriff ein");
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
        /// Called when the text in the search input has changed; displays new suggestions in the list view
        /// </summary>
        private void searchInput_TextChanged(object sender, TextChangedEventArgs e)
        {
            suggestionList.DisplaySearchSuggestions(searchInput.Text);
        }

        /// <summary>
        /// called when the ok button was pressed; opens the selected folder
        /// </summary>
        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            if (suggestionList.SelectedItem == null)
                return;
            int selectedId = (int)((ListViewItem)suggestionList.SelectedItem).Tag;
            var selectedFolder = ThisAddIn.GetFolderFromPath(ThisAddIn.Structure.GetPath(ThisAddIn.DbConn, selectedId));
            if(selectedFolder == null)
            {
                MessageBox.Show("Der ausgewählte Ordner exestiert nicht mehr. Bitte synchronisieren Sie die Ordnerstruktur.",
                    "Ordner exestiert nicht mehr",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }
            
            Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder = selectedFolder;
            Close();
        }

        /// <summary>
        /// called when the cancel button was pressed; closes this window
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void suggestionList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            okButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        }
    }
}
