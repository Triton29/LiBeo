using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.Windows;
using System.Windows.Controls;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace LiBeo
{
    /// <summary>
    /// Represents a folder structure
    /// </summary>
    public class FolderStructure
    {
        public Outlook.Folder RootFolder { get; set; }

        /// <summary>
        /// Constructor of class FolderStructor
        /// </summary>
        /// <param name="_rootFolder">The root folder of the folder structure</param>
        public FolderStructure(Outlook.Folder _rootFolder)
        {
            RootFolder = _rootFolder;
        }

        /// <summary>
        /// Saves the folder structure to a database
        /// </summary>
        /// <param name="conn">The SQLite database connection</param>
        public void SaveToDB(SQLiteConnection conn)
        {
            SQLiteCommand cmd = new SQLiteCommand(conn);

            // create a table if it does not exist
            cmd.CommandText = 
                "CREATE TABLE IF NOT EXISTS folders (" +
                "name varchar(255), id INTEGER PRIMARY KEY AUTOINCREMENT, parent_id int, got_deleted bit, UNIQUE(name, parent_id))";
            cmd.ExecuteNonQuery();

            // prepare for delete check
            cmd.CommandText = "UPDATE folders SET got_deleted=1";
            cmd.ExecuteNonQuery();

            // insert the root folder
            cmd.CommandText = "INSERT OR IGNORE INTO folders (name, parent_id, got_deleted) VALUES ('root', 0, 0)";
            cmd.ExecuteNonQuery();

            // root folder cannot be deleted
            cmd.CommandText = "UPDATE folders SET got_deleted=0 WHERE id=1";
            cmd.ExecuteNonQuery();

            InsertChildFolders(cmd, RootFolder, 1);

            // delete all deleted folders
            cmd.CommandText = "DELETE FROM folders WHERE got_deleted=1";
            cmd.ExecuteNonQuery();
        }

        /// <summary>
        /// Inserts all child folders of a folder into a database; recursive function
        /// </summary>
        /// <param name="cmd">SQLite command with database connection</param>
        /// <param name="parentFolder">The parent folder of the current child folders</param>
        /// <param name="parentId">The database-id of the parent folder</param>
        private void InsertChildFolders(SQLiteCommand cmd, Outlook.Folder parentFolder, int parentId)
        {
            foreach(Outlook.Folder folder in parentFolder.Folders)
            {
                // insert folder if not already inserted
                cmd.CommandText = "INSERT OR IGNORE INTO folders (name, parent_id, got_deleted) VALUES (@name, @parent_id, 0) ";
                cmd.Parameters.AddWithValue("@name", folder.Name);
                cmd.Parameters.AddWithValue("@parent_id", parentId);
                cmd.Prepare();
                cmd.ExecuteNonQuery();

                // get autoincremented id from current folder
                cmd.CommandText = "SELECT id FROM folders WHERE name=@name AND parent_id=@parent_id";
                cmd.Parameters.AddWithValue("@name", folder.Name);
                cmd.Parameters.AddWithValue("@parent_id", parentId);
                cmd.Prepare();
                SQLiteDataReader dataReader = cmd.ExecuteReader();
                dataReader.Read();
                int id = dataReader.GetInt32(0);
                dataReader.Close();

                // confirm that folder did not get deleted
                cmd.CommandText = "UPDATE folders SET got_deleted=0 WHERE id=@id";
                cmd.Parameters.AddWithValue("@id", id);
                cmd.Prepare();
                cmd.ExecuteNonQuery();

                // reset auto increment
                cmd.CommandText = "DELETE FROM sqlite_sequence WHERE name='folders'";
                cmd.ExecuteNonQuery();

                // revise all for the child folder
                InsertChildFolders(cmd, folder, id);
            }
        }

        /// <summary>
        /// Displays the folder structure from a database in a tree view
        /// </summary>
        /// <param name="conn">The SQLite database connection</param>
        /// <param name="treeView">The treeview in which the structure will be displayed</param>
        /// <param name="rootName">The name of the root folder</param>
        /// <param name="createCheckBoxes">If checkboxes before the header should be created; excelent for multi-select</param>
        public void DisplayInTreeView(SQLiteConnection conn, TreeView treeView, string rootName, bool createCheckBoxes)
        {
            TreeViewItem item = new TreeViewItem() { Header = rootName, Tag = 1, IsExpanded = true };
            treeView.Items.Add(item);
            AddChildItems(conn, item, 1, createCheckBoxes);
        }

        /// <summary>
        /// Adds child folders from a parent folder from a database to a treeview; recursive function
        /// </summary>
        /// <param name="conn">The SQLite database connection</param>
        /// <param name="parentItem">The parent-TreeViewItem from the child folders</param>
        /// <param name="parentId">The id from the parent folder in the database</param>
        /// <param name="createCheckBoxes">If checkboxes before the header should be created; excelent for multi-select</param>
        private void AddChildItems(SQLiteConnection conn, TreeViewItem parentItem, int parentId, bool createCheckBoxes)
        {
            SQLiteCommand cmd = new SQLiteCommand(conn);
            cmd.CommandText = "SELECT * FROM folders WHERE parent_id=@id ORDER BY name ASC";
            cmd.Parameters.AddWithValue("@id", parentId);
            cmd.Prepare();
            SQLiteDataReader dataReader = cmd.ExecuteReader();

            while (dataReader.Read())
            {
                TreeViewItem childItem;
                if (createCheckBoxes)
                {
                    WrapPanel wrapPanel = new WrapPanel();
                    wrapPanel.Children.Add(new CheckBox() { Content = dataReader.GetString(0), Tag = dataReader.GetInt32(1) });
                    childItem = new TreeViewItem() { Header = wrapPanel };
                    parentItem.Items.Add(childItem);
                }
                else
                {
                    childItem = new TreeViewItem() { Header = dataReader.GetString(0), Tag = dataReader.GetInt32(1) };
                    parentItem.Items.Add(childItem);
                }
                AddChildItems(conn, childItem, dataReader.GetInt32(1), createCheckBoxes);
            }
            dataReader.Close();
        }

        /// <summary>
        /// Gets the path of a folder in a folder structure saved in a database
        /// </summary>
        /// <param name="conn">SQLite database connection</param>
        /// <param name="folderId">The id of the folder in the database</param>
        /// <returns>The path of the folder in a list</returns>
        public static List<string> GetPath(SQLiteConnection conn, int folderId)
        {
            List<string> path = new List<string>();
            int parentId = folderId;
            SQLiteCommand cmd = new SQLiteCommand(conn);

            while (parentId != 1)
            {
                cmd.CommandText = "SELECT name, parent_id FROM folders WHERE id=@id";
                cmd.Parameters.AddWithValue("@id", parentId);
                cmd.Prepare();
                SQLiteDataReader dataReader = cmd.ExecuteReader();
                dataReader.Read();

                path.Insert(0, dataReader.GetString(0));
                parentId = dataReader.GetInt32(1);

                dataReader.Close();
            }

            return path;
        }
    }
}
