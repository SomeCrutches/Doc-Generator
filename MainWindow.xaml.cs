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
using System.Windows.Xps.Packaging;




namespace ProizvodkaWPF
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            WindowState = WindowState.Maximized;
            string gesturefile = System.IO.Path.Combine(Environment.CurrentDirectory, @"Docs\Восстановление-образец.xps");
            if (File.Exists(gesturefile))
            {
                XpsDocument doc_pre = new XpsDocument(gesturefile, FileAccess.Read);
                FirstTab.Document = doc_pre.GetFixedDocumentSequence();
            }
            else
            {
                MessageBox.Show("Oops, ne workaet");
            }

           
        }
        
        private void New_tab_button(object sender, RoutedEventArgs e)
        {

            PreviewDocument preview_doc = new PreviewDocument();
            preview_doc.Owner = this;
            preview_doc.Doc_type = "Hello";
            preview_doc.Show();
        }

        private void Test_Click(object sender, RoutedEventArgs e)
        {
            var filename = System.IO.Path.Combine(Environment.CurrentDirectory, @"Docs\Восстановление-образец.docx");
            var paster = new WordPaster(filename);

            //Указываем ключи
            var items = new Dictionary<string, string>
            {
                { "<STUDENT_FIO>", FirstTabFirst.Text },
                { "<NAPRAVLENIE_PODGOTOVKI>", FirstTabSecond.Text},
                { "<NAPRAVLENIE_PROGRAMMI>", FirstTabThird.Text },
                { "<DOGOVOR_NUMBER>", FirstTabFourth.Text },
                { "<STUDENT_COURSE>", FirstTabFifth.Text },
                { "<PRIKAZ_OTCHISL>", FirstTabSixth.Text }
               
            };
            paster.Process(items);
        }
    }
}
