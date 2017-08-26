using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using Hardcodet.Wpf.TaskbarNotification;
using VocabunaryReminder.Properties;

namespace VocabunaryReminder
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public TaskbarIcon notifyIcon;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            notifyIcon = new TaskbarIcon();


            //ExcelReader excelReader = new ExcelReader();
            //excelReader.GenerateEXcelData("側");
            // base.OnLoad(e);
            notifyIcon.Icon = VocabunaryReminder.Properties.Resources.Led;
            notifyIcon.ToolTipText = "Left-click to open popup";
            notifyIcon.Visibility = Visibility.Visible;

            notifyIcon.TrayPopup = new VocabunaryNofification();

        }
    }
}
