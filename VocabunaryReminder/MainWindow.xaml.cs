using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
namespace VocabunaryReminder
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            ExcelReader excelReader = new ExcelReader();
            excelReader.GenerateEXcelData("側");


        }
    }
}
