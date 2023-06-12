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
using System.Windows.Xps;



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
        }

        private void UploadButton_click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "XPS Files (*.xps)|*.xps";
            bool? response = openFileDialog.ShowDialog();
            if (response == true)
            {
                string filename = openFileDialog.FileName;
                /*
                 * 
                XpsDocument doc = new XpsDocument(filename, FileAccess.Write);
                XpsDocumentWriter writer = XpsDocument.CreateXpsDocumentWriter(doc);
                writer.Write(DocView.Document as FixedDocument);
                doc.Close();
                //MessageBox.Show(filename);
                */
                XpsDocument doc = new XpsDocument(filename, FileAccess.Read);
                DocView.Document = doc.GetFixedDocumentSequence();
            }
        }
    }
}
