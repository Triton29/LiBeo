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

        /// <summary>
        /// manage first focused elements
        /// </summary>
        protected override void OnContentRendered(EventArgs e)
        {
            base.OnContentRendered(e);

            if (tabConrol.SelectedIndex == 1)
                Keyboard.Focus(folderExplorer);
            if (tabConrol.SelectedIndex == 2)
                Keyboard.Focus(quickAccessList);
        }

        public Actions()
        {
            InitializeComponent();

            // display folder structure
            ThisAddIn.DbConn.Open();
            ThisAddIn.Structure.DisplayInTreeView(ThisAddIn.DbConn, folderExplorer, ThisAddIn.Name, false);
            ThisAddIn.DbConn.Close();

            // display quick access list
            DisplayQuickAccessList(quickAccessList);
            if (quickAccessList.Items.Count == 0)
            {
                listViewEmptyLabel.Content = "Keine Elemente in der Schenllzugriffsliste";
            }
        }

        /// <summary>
        /// Called when OK-button is clicked; moves mail(s) to folder, selected in TreeView
        /// </summary>
        private void OKButton_Click(object sender, RoutedEventArgs e)
        {
            if(tabConrol.SelectedIndex == 0)    // automatic sort
            {

            }
            if(tabConrol.SelectedIndex == 1)    // manual sort
            {
                ThisAddIn.DbConn.Open();
                try
                {
                    TreeViewItem selectedItem = (TreeViewItem)folderExplorer.SelectedItem;
                    if (selectedItem == null)
                    {
                        ThisAddIn.DbConn.Close();
                        return;
                    }

                    int id = (int)selectedItem.Tag;

                    List<string> path = FolderStructure.GetPath(ThisAddIn.DbConn, id);
                    Outlook.Folder currentFolder = rootFolder;
                    foreach (string folder in path)
                    {
                        currentFolder = (Outlook.Folder)currentFolder.Folders[folder];
                    }
                    foreach (Outlook.MailItem mail in ThisAddIn.GetSelectedMails())
                    {
                        mail.Move(currentFolder);
                    }

                    ThisAddIn.DbConn.Close();
                    this.Close();
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    MessageBox.Show("Der ausgewählte Ordner exestiert nicht mehr. Bitte synchronisieren Sie die Ordnerstruktur.",
                        "Ordner exestiert nicht mehr",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
                ThisAddIn.DbConn.Close();
            }
            if(tabConrol.SelectedIndex == 2)    // quick access list sort
            {
                ThisAddIn.DbConn.Open();
                int id = (int)((ListViewItem)quickAccessList.SelectedItem).Tag;

                List<string> path = FolderStructure.GetPath(ThisAddIn.DbConn, id);
                Outlook.Folder currentFolder = rootFolder;
                foreach (string folder in path)
                {
                    currentFolder = (Outlook.Folder)currentFolder.Folders[folder];
                }
                foreach (Outlook.MailItem mail in ThisAddIn.GetSelectedMails())
                {
                    mail.Move(currentFolder);
                }

                this.Close();
                ThisAddIn.DbConn.Close();
            }
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

        /// <summary>
        /// Called when the user selects another tab; calls the MoveToTray method when the 4th tab is selected
        /// </summary>
        private void tabConrol_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if(tabConrol.SelectedIndex == 3)
            {
                MoveToTray();
                this.Close();
            }
        }

        /// <summary>
        /// Moves the selected mails to a user-defined tray
        /// </summary>
        public static void MoveToTray()
        {
            string outgoingFolder = "Gesendet";
            string incomingFolder = "Empfangen";

            Outlook.Folder trayFolder;
            try     // check if tray path exists
            {
                trayFolder = ThisAddIn.GetFolderFromPath(Properties.Settings.Default.TrayPath);
            }
            catch
            {
                MessageBox.Show("Der Ablage-Ordner exestiert nicht! Ändern Sie ihn in den Add-In-Einstellungen.", 
                    "Ablage-Ordner nicht gefunden", 
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }

            Outlook.Folder currentFolder = trayFolder;
            foreach(Outlook.MailItem mail in ThisAddIn.GetSelectedMails())
            {
                if(mail.SenderEmailAddress.ToLower() == ThisAddIn.EmailAddress.ToLower() &&
                    mail.SenderEmailAddress.ToLower() != mail.Recipients[1].Address.ToLower())   // if mail is outgoing
                {
                    if(IsInFolder(currentFolder, mail.CreationTime.Year.ToString()))
                        currentFolder = (Outlook.Folder)currentFolder.Folders[mail.CreationTime.Year.ToString()];
                    else
                        currentFolder = (Outlook.Folder)currentFolder.Folders.Add(mail.CreationTime.Year.ToString());

                    if (IsInFolder(currentFolder, outgoingFolder))
                        currentFolder = (Outlook.Folder)currentFolder.Folders[outgoingFolder];
                    else
                        currentFolder = (Outlook.Folder)currentFolder.Folders.Add(outgoingFolder);

                    try
                    {
                        mail.Move(currentFolder);
                    }
                    catch
                    {
                        return;
                    }
                }
                else    // if mail is incoming
                {
                    if (IsInFolder(currentFolder, mail.ReceivedTime.Year.ToString()))
                        currentFolder = (Outlook.Folder)currentFolder.Folders[mail.ReceivedTime.Year.ToString()];
                    else
                        currentFolder = (Outlook.Folder)currentFolder.Folders.Add(mail.ReceivedTime.Year.ToString());

                    if (IsInFolder(currentFolder, incomingFolder))
                        currentFolder = (Outlook.Folder)currentFolder.Folders[incomingFolder];
                    else
                        currentFolder = (Outlook.Folder)currentFolder.Folders.Add(incomingFolder);

                    try
                    {
                        mail.Move(currentFolder);
                    }
                    catch
                    {
                        return;
                    }
                }
                currentFolder = trayFolder;
            }
        }

        /// <summary>
        /// Checks if a subfolder in an outlook folder exists
        /// </summary>
        /// <param name="folder">The parent folder</param>
        /// <param name="subFolderToCheck">The name of the subfolder to check</param>
        /// <returns>true if subfolder exists; false if not</returns>
        public static bool IsInFolder(Outlook.Folder folder, string subFolderToCheck)
        {
            foreach(Outlook.Folder f in folder.Folders)
            {
                if (f.Name == subFolderToCheck)
                    return true;
            }
            return false;
        }

        /// <summary>
        /// Displays a the quick access list of folders saved in the database in a list view
        /// </summary>
        /// <param name="list">The list view where the folders are displayed</param>
        public static void DisplayQuickAccessList(ListView list)
        {
            ThisAddIn.DbConn.Open();
            SQLiteCommand dbCmd = new SQLiteCommand(ThisAddIn.DbConn);

            dbCmd.CommandText = "SELECT * FROM quick_access_folders";
            SQLiteDataReader dataReader = dbCmd.ExecuteReader();
            while (dataReader.Read())
            {
                int id = dataReader.GetInt32(0);
                List<string> path = FolderStructure.GetPath(ThisAddIn.DbConn, id);
                ListViewItem item = new ListViewItem() { Content = path[path.Count - 1], Tag = id};
                list.Items.Add(item);
            }
            ThisAddIn.DbConn.Close();
        }
    }
}
