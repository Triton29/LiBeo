﻿using System;
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
    /// Interaction logic for MultiSelectFolder.xaml
    /// </summary>
    public partial class MultiSelectFolder : Window
    {
        public List<int> PreSelectedFolderIds = new List<int>();
        public HashSet<int> SelectedFolderIds = new HashSet<int>();
        public bool Canceled = true;

        /// <summary>
        /// Checks pre-selected folders
        /// </summary>
        protected override void OnContentRendered(EventArgs e)
        {
            base.OnContentRendered(e);

            // load current setting
            LoadCurrentSetting(PreSelectedFolderIds);
        }

        public MultiSelectFolder()
        {
            InitializeComponent();

            // display folder structure in tree view
            ThisAddIn.Structure.DisplayInTreeView(ThisAddIn.DbConn, folderExplorer, ThisAddIn.Name, true);
        }

        /// <summary>
        /// Called when the ok button is pressed; 
        /// checks all check boxes and writes the folder paths of the checked checkboxes in the SelectedFolderPaths list;
        /// sets the Canceled property to false and closes the window
        /// </summary>
        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedSuggestedItems = searchSuggestions.SelectedItems;
            if(selectedSuggestedItems != null)
            {
                foreach(var item in selectedSuggestedItems)
                {
                    SelectedFolderIds.Add((int) ((ListViewItem)item).Tag);
                }
            }
            
            foreach (CheckBox checkBox in ThisAddIn.GetLogicalChildren<CheckBox>(folderExplorer))
            {
                if(checkBox.IsChecked == true)
                {
                    int id = (int) checkBox.Tag;
                    SelectedFolderIds.Add(id);
                }
            }

            if(SelectedFolderIds.Count > 0)
            {
                Canceled = false;
                this.Close();
            }
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
            this.Close();
        }

        /// <summary>
        /// Loads the current folder setting into the tree view
        /// </summary>
        /// <param name="selectedFolderIds">A list with the folder ids</param>
        public void LoadCurrentSetting(List<int> selectedFolderIds)
        {
            foreach (CheckBox checkBox in ThisAddIn.GetLogicalChildren<CheckBox>(folderExplorer))
            {
                int id = (int)checkBox.Tag;
                
                if (selectedFolderIds.Contains(id))
                {
                    checkBox.IsChecked = true;
                }
            }
        }

        private void searchInput_TextChanged(object sender, TextChangedEventArgs e)
        {
            searchSuggestions.DisplaySearchSuggestions(searchInput.Text);
            folderExplorer.Visibility = searchInput.Text == "" ? Visibility.Visible : Visibility.Collapsed;
            searchSuggestions.Visibility = searchInput.Text == "" ? Visibility.Collapsed : Visibility.Visible;
        }

        private void searchSuggestions_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {

        }
    }
}
