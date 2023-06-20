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
using Microsoft.WindowsAPICodePack.Dialogs;




namespace ProizvodkaWPF
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string folderName;
        string defaultpath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Документы дирекция КГЭУ");

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

            if (Properties.Settings.Default.save_path == defaultpath)
            {
                PathLabelUpdate("Рабочий стол, папка 'Документы дирекция КГЭУ'");
            }
            else
            {
                PathLabelUpdate(Properties.Settings.Default.save_path);
            }

        }
        
        private void New_tab_button(object sender, RoutedEventArgs e)
        {

            PreviewDocument preview_doc = new PreviewDocument();
            preview_doc.Owner = this;
            preview_doc.Doc_type = "Hello";
            preview_doc.Show();
        }

        private void Test_Click(object sender, RoutedEventArgs e)   // old code
        {
            var filename = System.IO.Path.Combine(Environment.CurrentDirectory, @"Docs\Восстановление-образец.docx");
            var paster = new WordPaster(filename);

            //Указываем ключи
            var items = new Dictionary<string, string>
            {
                { "<STUDENT_FIO>", FirstTab1.Text },
                { "<NAPRAVLENIE_PODGOTOVKI>", FirstTab2.Text},
                { "<NAPRAVLENIE_PROGRAMMI>", FirstTab3.Text },
                { "<DOGOVOR_NUMBER>", FirstTab4.Text },
                { "<STUDENT_COURSE>", FirstTab5.Text },
                { "<PRIKAZ_OTCHISL>", FirstTab6.Text },
                { "<ELIMINATE_DAY>", FirstTab7.Text },
                { "<NEW_GROUP_DATE>", FirstTab8.Text },
                { "<PROTOCOL_ZASED>", FirstTab9.Text },
                { "<DOG_PLATN>", FirstTab10.Text }

            };
            paster.Process(items);
        }

        private void FrstTbGen(object sender, RoutedEventArgs e)
        {
            var filename = System.IO.Path.Combine(Environment.CurrentDirectory, @"Docs\Восстановление-образец.docx");
            var paster = new WordPaster(filename);

            //Указываем ключи
            var items = new Dictionary<string, string>
            {
                { "<STUDENT_FIO>", FirstTab1.Text },
                { "<NAPRAVLENIE_PODGOTOVKI>", FirstTab2.Text},
                { "<NAPRAVLENIE_PROGRAMMI>", FirstTab3.Text },
                { "<DOGOVOR_NUMBER>", FirstTab4.Text },
                { "<STUDENT_COURSE>", FirstTab5.Text },
                { "<PRIKAZ_OTCHISL>", FirstTab6.Text },
                { "<ELIMINATE_DAY>", FirstTab7.Text },
                { "<NEW_GROUP_DATE>", FirstTab8.Text },
                { "<PROTOCOL_ZASED>", FirstTab9.Text },
                { "<DOG_PLATN>", FirstTab10.Text }

            };
            paster.Process(items);
        }




        

        private void PathLabelUpdate(string value)
        {
            Save_Path_Label.Content = value;
        }

        

        private void Set_Default_Path_Button_Click(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.save_path = defaultpath;
            Properties.Settings.Default.Save();
            MessageBox.Show(
            "Установлен путь для сохранения: Рабочий стол, папка 'Документы дирекция КГЭУ'",
            "Успешно");
            PathLabelUpdate("Рабочий стол, папка 'Документы дирекция КГЭУ'");
        }

        private void Change_Save_Path_Button_Click(object sender, RoutedEventArgs e)
        {
            //показать диалог выбора папки
            var dlg = new CommonOpenFileDialog
            {
                Title = "Выбор папки",
                IsFolderPicker = true,
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),

                AddToMostRecentlyUsedList = false,
                AllowNonFileSystemItems = false,
                DefaultDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                EnsureFileExists = true,
                EnsurePathExists = true,
                EnsureReadOnly = false,
                EnsureValidNames = true,
                Multiselect = false,
                ShowPlacesList = true
            };

            //если папка выбрана и нажата клавиша `OK` - значит можно получить путь к папке
            if (dlg.ShowDialog() == CommonFileDialogResult.Ok)
            {
                //запишем в нашу переменную путь к папке
                folderName = dlg.FileName;
                
            }
            //записываем в параметр значение переменной и сохраняем
            Properties.Settings.Default.save_path = folderName;
            Properties.Settings.Default.Save();
            MessageBox.Show(
            "Установлен путь для сохранения: " + folderName,
            "Успешно");
            PathLabelUpdate(Properties.Settings.Default.save_path);
        }
    }
}
