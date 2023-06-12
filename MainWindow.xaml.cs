using System;
using System.Collections.Generic;
using System.IO.Packaging;
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
using System.IO;




namespace ProizvodkaWPF
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            WindowState = WindowState.Maximized;

            InitializeComponent();
            
        }

        private void New_tab_button(object sender, RoutedEventArgs e)
        {

            PreviewDocument preview_doc = new PreviewDocument();
            preview_doc.Owner = this;
            preview_doc.Doc_type = "Hello";
            preview_doc.Show();
        }
    }
}
