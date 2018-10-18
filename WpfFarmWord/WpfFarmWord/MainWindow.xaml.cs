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

namespace WpfFarmWord
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string filepath =
@"C:\Users\1\source\repos\WpfFarmWord\WpfFarmWord\bin\Debug\004.docx";
            //@"C:\Users\1\Documents\Visual Studio 2012\Projects\WpfFarmingWord\WpfFarmingWord\bin\Debug\004.docx";
            FarmWord fw = new FarmWord(filepath);
        }
    }
}
