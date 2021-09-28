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
using System.Windows.Media.Animation;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace LiBeo
{
    /// <summary>
    /// Interaction logic for WaitWindow.xaml
    /// </summary>
    public partial class WaitWindow : Window
    {
        string loadingSource = AppDomain.CurrentDomain.BaseDirectory + @"img\loading.gif";
        public WaitWindow()
        {
            InitializeComponent();
            WpfAnimatedGif.ImageBehavior.SetAnimatedSource(loading, new BitmapImage(new Uri(loadingSource)));
        }
    }
}
