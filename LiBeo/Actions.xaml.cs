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
using Office = Microsoft.Office.Core;
using System.Data.SQLite;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace LiBeo
{
    /// <summary>
    /// Interaction logic for Actions.xaml
    /// </summary>
    public partial class Actions : Window
    {
        private Outlook.Folder rootFolder = ThisAddIn.RootFolder;
        public Actions()
        {
            InitializeComponent();

            SQLiteConnection dbConn = new SQLiteConnection("Data Source=" + ThisAddIn.DbPath);
            dbConn.Open();
            ThisAddIn.Structure.DisplayInTreeView(dbConn, folderExplorer);
            dbConn.Close();
        }

        /// <summary>
        /// Called when OK-button is clicked; moves mail(s) to folder, selected in TreeView
        /// </summary>
        private void OKButton_Click(object sender, RoutedEventArgs e)
        {
            SQLiteConnection dbConn = new SQLiteConnection("Data Source=" + ThisAddIn.DbPath);
            dbConn.Open();
            int id = (int)((TreeViewItem)folderExplorer.SelectedItem).Tag;
            List<string> path = FolderStructure.GetPath(dbConn, id);
            Outlook.Folder currentFolder = rootFolder;
            foreach(string folder in path)
            {
                currentFolder = (Outlook.Folder) currentFolder.Folders[folder];
            }
            foreach(Outlook.MailItem mail in ThisAddIn.GetSelectedMails())
            {
                mail.Move(currentFolder);
            }

            dbConn.Close();
            this.Close();
        }

        /// <summary>
        /// Called when any key on the keyboard is pressed; implements shortcuts
        /// </summary>
        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Enter)
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
    }
}
