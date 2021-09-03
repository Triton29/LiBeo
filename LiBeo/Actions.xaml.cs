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

            if (tabConrol.SelectedIndex == 0)
                Keyboard.Focus(autoSortList);
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
                quickAccessListEmpty.Content = "Keine Elemente in der Schenllzugriffsliste";
            }

            IEnumerable<Outlook.MailItem> selectedMails;
            try
            {
                selectedMails = ThisAddIn.GetSelectedMails();
            }
            catch
            {
                this.Close();
                MessageBox.Show("Bitte wählen Sie nur E-Mails aus",
                        "Ausgewählte Elemente sind keine E-Mails",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                return;
            }
            
            if(selectedMails.Count() == 1)
            {
                DisplayAutoSortList(autoSortList, selectedMails.First());
                if(autoSortList.Items.Count == 0)
                {
                    autoSortListEmpty.Content = "Keine Vorschläge für diese E-Mail gefunden";
                }
            }
            else
            {
                autoSortListEmpty.Content = "Keine Vorschläge gefunden, da mehrere bzw. keine E-Mails ausgewählt wurden";
            }
        }

        /// <summary>
        /// Gets an outlook folder from the folder path
        /// </summary>
        /// <param name="path">The folder path as a list</param>
        /// <returns>The wanted outlook folder</returns>
        static Outlook.Folder GetFolderFromPath(List<string> path)
        {
            Outlook.Folder folder = ThisAddIn.RootFolder;
            foreach(string folderName in path)
            {
                folder = (Outlook.Folder)folder.Folders[folderName];
            }
            return folder;
        }

        /// <summary>
        /// Moves a collection of mails (or one mail) into antoher folder
        /// </summary>
        /// <param name="mails">The collection of mails to move</param>
        /// <param name="folderId">The id of the folder in the database</param>
        static void MoveMails(IEnumerable<Outlook.MailItem> mails, int folderId)
        {
            List<string> path = FolderStructure.GetPath(ThisAddIn.DbConn, folderId);
            Outlook.Folder targetFolder = GetFolderFromPath(path);
            foreach (Outlook.MailItem mail in mails)
            {
                mail.Move(targetFolder);
            }
        }

        /// <summary>
        /// Called when OK-button is clicked; moves mail(s) to folder, selected in TreeView
        /// </summary>
        private void OKButton_Click(object sender, RoutedEventArgs e)
        {
            ThisAddIn.DbConn.Open();
            int id = -1;
            if (tabConrol.SelectedIndex == 0)    // automatic sort
            {
                try
                {
                    ListViewItem selectedItem = (ListViewItem)autoSortList.SelectedItem;
                    if (selectedItem == null)
                    {
                        ThisAddIn.DbConn.Close();
                        return;
                    }

                    id = (int)selectedItem.Tag;

                    MoveMails(ThisAddIn.GetSelectedMails(), id);

                    this.Close();
                }
                catch
                {
                    MessageBox.Show("Der ausgewählte Ordner exestiert nicht mehr. Bitte synchronisieren Sie die Ordnerstruktur.",
                        "Ordner exestiert nicht mehr",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
            }
            if(tabConrol.SelectedIndex == 1)    // manual sort
            {
                try
                {
                    TreeViewItem selectedItem = (TreeViewItem)folderExplorer.SelectedItem;
                    if (selectedItem == null)
                    {
                        ThisAddIn.DbConn.Close();
                        return;
                    }

                    id = (int)selectedItem.Tag;

                    MoveMails(ThisAddIn.GetSelectedMails(), id);

                    this.Close();
                }
                catch
                {
                    MessageBox.Show("Der ausgewählte Ordner exestiert nicht mehr. Bitte synchronisieren Sie die Ordnerstruktur.",
                        "Ordner exestiert nicht mehr",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
            }
            if(tabConrol.SelectedIndex == 2)    // quick access list sort
            {
                try
                {
                    ListViewItem selectedItem = (ListViewItem)quickAccessList.SelectedItem;
                    if (selectedItem == null)
                    {
                        ThisAddIn.DbConn.Close();
                        return;
                    }

                    id = (int)selectedItem.Tag;

                    MoveMails(ThisAddIn.GetSelectedMails(), id);

                    this.Close();
                }
                catch
                {
                    MessageBox.Show("Der ausgewählte Ordner exestiert nicht mehr. Bitte synchronisieren Sie die Ordnerstruktur.",
                        "Ordner exestiert nicht mehr",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
            }

            if(id != -1 && learnNothingCheckBox.IsChecked == false)
            {
                foreach(Outlook.MailItem mail in ThisAddIn.GetSelectedMails())
                    LearnTags(mail.Subject, id);
            }

            ThisAddIn.DbConn.Close();
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
        /// Saves a string (subject of a mail) and the mail's target folder id in a database in a table named current_mail_subject;
        /// every word has an entry; stop words (from the stop_words table) will be deleted
        /// </summary>
        /// <param name="dbCmd">SQLiteCommand of the database</param>
        /// <param name="subject">The string which should be saved</param>
        /// <param name="targetFolder">The mail's target folder id</param>
        static void SubjectToDb(SQLiteCommand dbCmd, string subject, int targetFolder)
        {
            dbCmd.CommandText = "DELETE FROM current_mail_subject";
            dbCmd.ExecuteNonQuery();

            foreach (string word in subject.Split(' '))
            {
                string rawWord = string.Concat(word.Where(char.IsLetterOrDigit));
                if (rawWord != "")
                {
                    dbCmd.CommandText = "INSERT OR IGNORE INTO current_mail_subject (folder, word) VALUES (@id, @word)";
                    dbCmd.Parameters.AddWithValue("@id", targetFolder);
                    dbCmd.Parameters.AddWithValue("@word", rawWord);
                    dbCmd.Prepare();
                    dbCmd.ExecuteNonQuery();
                }
                dbCmd.CommandText = "DELETE FROM current_mail_subject WHERE word IN (SELECT word FROM stop_words)";
                dbCmd.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// Learns tags for mail-folders and saves them in a database
        /// </summary>
        /// <param name="subject">The subject of the mail</param>
        /// <param name="folderId">The folder id of the folder in which the mail was moved</param>
        public static void LearnTags(string subject, int folderId)
        {
            SQLiteCommand dbCmd = new SQLiteCommand(ThisAddIn.DbConn);

            SubjectToDb(dbCmd, subject, folderId);

            dbCmd.CommandText = "INSERT OR IGNORE INTO tags (folder, tag) SELECT folder, word FROM current_mail_subject";
            dbCmd.ExecuteNonQuery();
        }

        /// <summary>
        /// Displays folder suggestions for a mail in a ListView (AutoSortList)
        /// </summary>
        /// <param name="list">The ListView</param>
        /// <param name="mail">The mail</param>
        public static void DisplayAutoSortList(ListView list, Outlook.MailItem mail)
        {
            ThisAddIn.DbConn.Open();
            SQLiteCommand dbCmd = new SQLiteCommand(ThisAddIn.DbConn);

            SubjectToDb(dbCmd, mail.Subject, 0);
            dbCmd.CommandText = "SELECT folder FROM tags WHERE tag IN (SELECT word FROM current_mail_subject)";
            SQLiteDataReader dataReader = dbCmd.ExecuteReader();

            List<FolderSuggestion> folderSuggestions = new List<FolderSuggestion>();
            while (dataReader.Read())
            {
                FolderSuggestion folderSuggestion = folderSuggestions.Find(x => x.FolderId == dataReader.GetInt32(0));
                if(folderSuggestion == null)
                {
                    folderSuggestions.Add(new FolderSuggestion { FolderId = dataReader.GetInt32(0), Importance = 1 });
                }
                else
                {
                    folderSuggestion.Importance++;
                }
            }
            List<FolderSuggestion> sortedFolderSuggestions = folderSuggestions.OrderByDescending(x => x.Importance).ToList();
            foreach(FolderSuggestion suggestion in sortedFolderSuggestions)
            {
                var path = FolderStructure.GetPath(ThisAddIn.DbConn, suggestion.FolderId);
                ListViewItem item = new ListViewItem { Content = path[path.Count - 1], Tag = suggestion.FolderId };
                list.Items.Add(item);
            }

            ThisAddIn.DbConn.Close();
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

        /// <summary>
        /// Moves a folder into another folder, in outlook as well as in the database; the id stays the same, so the tags won't be lost
        /// </summary>
        public static void MoveFolder()
        {
            MultiSelectFolder foldersToMoveWindow = new MultiSelectFolder() { Title = "Ordner zum Verschieben auswählen" };
            if(foldersToMoveWindow.ShowDialog() == false && !foldersToMoveWindow.Canceled)
            {
                SelectFolder targetFolderWindow = new SelectFolder() { Title = "Ordner auswählen, in den der/die Ordner verschoben werden" };
                if(targetFolderWindow.ShowDialog() == false && !targetFolderWindow.Canceled)
                {
                    ThisAddIn.DbConn.Open();
                    SQLiteCommand dbCmd = new SQLiteCommand(ThisAddIn.DbConn);
                    foreach (int folderToMoveId in foldersToMoveWindow.SelectedFolderIds)
                    {
                        Outlook.Folder folderToMove = GetFolderFromPath(FolderStructure.GetPath(ThisAddIn.DbConn, folderToMoveId));
                        Outlook.Folder targetFolder = GetFolderFromPath(targetFolderWindow.SelectedFolderPath);
                        folderToMove.MoveTo(targetFolder);

                        dbCmd.CommandText = "UPDATE folders SET parent_id=@parent_id WHERE id=@id";
                        dbCmd.Parameters.AddWithValue("@parent_id", targetFolderWindow.SelectedFolderId);
                        dbCmd.Parameters.AddWithValue("@id", folderToMoveId);
                        dbCmd.ExecuteNonQuery();
                    }
                    ThisAddIn.DbConn.Close();
                }
            }
        }
    }

    /// <summary>
    /// Represents a folder suggestion for the AutoSort function
    /// </summary>
    public class FolderSuggestion
    {
        public int FolderId { get; set; }
        public int Importance { get; set; }
    }
}
