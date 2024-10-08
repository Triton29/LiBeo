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
    /// Interaction logic for SelectFolder.xaml
    /// </summary>
    public partial class SelectFolder : Window
    {
        public int SelectedFolderId = -1;
        public List<string> SelectedFolderPath = null;
        public bool Canceled = true;

        public SelectFolder()
        {
            InitializeComponent();

            SelectedFolderPath = new List<string>();

            // display folder structure in tree view
            ThisAddIn.Structure.DisplayInTreeView(ThisAddIn.DbConn, folderExplorer, ThisAddIn.Name, false);
        }

        /// <summary>
        /// called when the ok button is pressed; sets the canceled property to false and closes the window
        /// </summary>
        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            if(searchSuggestions.SelectedItem != null)
            {
                SelectedFolderId = (int)((ListViewItem)searchSuggestions.SelectedItem).Tag;
                SelectedFolderPath = ThisAddIn.Structure.GetPath(ThisAddIn.DbConn, SelectedFolderId);
            }
            if(SelectedFolderId != -1 && SelectedFolderPath != null)
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
        /// Called when the selected item in the folder explorer treeview has changed; 
        /// writes the currently selected folder in the public variable
        /// </summary>
        private void folderExplorer_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            int id = (int)((TreeViewItem)folderExplorer.SelectedItem).Tag;
            SelectedFolderId = id;
            SelectedFolderPath = ThisAddIn.Structure.GetPath(ThisAddIn.DbConn, id);
        }

        private void searchInput_TextChanged(object sender, TextChangedEventArgs e)
        {
            searchSuggestions.DisplaySearchSuggestions(searchInput.Text);
            folderExplorer.Visibility = searchInput.Text == "" ? Visibility.Visible : Visibility.Collapsed;
            searchSuggestions.Visibility = searchInput.Text == "" ? Visibility.Collapsed : Visibility.Visible;
        }

        private void searchSuggestions_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            okButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        }
    }
}
