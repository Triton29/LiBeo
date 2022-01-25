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
    /// Interaction logic for FolderList.xaml
    /// </summary>
    public partial class FolderList : UserControl
    {
        public object SelectedItem { get; set; }

        public FolderList()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Displays folder search suggestions for a pattern
        /// </summary>
        /// <param name="pattern">Pattern of the search</param>
        public void DisplaySearchSuggestions(string pattern)
        {
            list.Items.Clear();
            if(pattern == "")
            {
                DisplayListViewMsg("Geben Sie einen Suchbegriff ein");
                return;
            }
            foreach (int id in ThisAddIn.Structure.SearchFolder(ThisAddIn.DbConn, pattern))
            {
                List<string> path = ThisAddIn.Structure.GetPath(ThisAddIn.DbConn, id);
                string pathStr = string.Join("\\", path);
                ListViewItem item = new ListViewItem()
                {
                    Content = pathStr,
                    Tag = id
                };
                list.Items.Add(item);
            }
            if (list.Items.Count == 0)
                DisplayListViewMsg("Keinen passenden Ordner gefunden");
        }

        /// <summary>
        /// Displays a message in the suggestion list view
        /// </summary>
        /// <param name="text">Message that should be displayed</param>
        public void DisplayListViewMsg(String text)
        {
            ListViewItem item = new ListViewItem()
            {
                Content = text,
                Foreground = Brushes.DarkGray,
                HorizontalAlignment = HorizontalAlignment.Center
            };
            list.Items.Add(item);
        }

        /// <summary>
        /// Called when list view selection has changed
        /// </summary>
        /// <param name="sender"></param>
        private void list_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            this.SelectedItem = list.SelectedItem;
        }
    }
}
