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
using System.Windows.Shapes;
using System.Windows.Xps.Packaging;
using System.Windows.Xps;
using System.IO.Packaging;
using System.Windows.Navigation;
using System.IO;

namespace ProizvodkaWPF
{
    /// <summary>
    /// Логика взаимодействия для PreviewDocument.xaml
    /// </summary>
    public partial class PreviewDocument : Window
    {
        public string Doc_type { get; set; }

        public PreviewDocument()
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

        private void UploadButtonAbra_click(object sender, RoutedEventArgs e)
        {
            string gesturefile = System.IO.Path.Combine(Environment.CurrentDirectory, @"Индивидуальное-задание_.xps");
            if (File.Exists(gesturefile))
            {
                XpsDocument doc_pre = new XpsDocument(gesturefile, FileAccess.Read);
                DocView.Document = doc_pre.GetFixedDocumentSequence();
            }
            else
            {
                MessageBox.Show("Oops, ne workaet");
            }
        }
    }
}
