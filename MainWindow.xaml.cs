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

            this.Title = "Генерация документов";
            Properties.Settings.Default.save_path = defaultpath;
            WindowState = WindowState.Maximized;

            string gesturefile = System.IO.Path.Combine(Environment.CurrentDirectory, @"Docs\Восстановление-образец.xps");
            if (File.Exists(gesturefile))
            {
                XpsDocument doc_pre = new XpsDocument(gesturefile, FileAccess.Read);
                FirstTab.Document = doc_pre.GetFixedDocumentSequence();
            }
            else
            {
                MessageBox.Show(".xps file doesn't exist");
            }

            string gesturefile1 = System.IO.Path.Combine(Environment.CurrentDirectory, @"Docs\отчисление-по-сб.ж-образец.xps");
            if (File.Exists(gesturefile1))
            {
                XpsDocument doc_pre = new XpsDocument(gesturefile1, FileAccess.Read);
                SecondTab.Document = doc_pre.GetFixedDocumentSequence();
            }
            else
            {
                MessageBox.Show(".xps file doesn't exist");
            }

            string gesturefile2 = System.IO.Path.Combine(Environment.CurrentDirectory, @"Docs\Перевод-с-ОП-на-другую-ОП-образец.xps");
            if (File.Exists(gesturefile2))
            {
                XpsDocument doc_pre = new XpsDocument(gesturefile2, FileAccess.Read);
                ThirdTab.Document = doc_pre.GetFixedDocumentSequence();
            }
            else
            {
                MessageBox.Show(".xps file doesn't exist");
            }

            string gesturefile3 = System.IO.Path.Combine(Environment.CurrentDirectory, @"Docs\приказ-о-зачислении-в-порядке-перевода-образец.xps");
            if (File.Exists(gesturefile3))
            {
                XpsDocument doc_pre = new XpsDocument(gesturefile3, FileAccess.Read);
                FourthTab.Document = doc_pre.GetFixedDocumentSequence();
            }
            else
            {
                MessageBox.Show(".xps file doesn't exist");
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

        /*
        private void Test_Click(object sender, RoutedEventArgs e)   // old code
        {
            var filename = System.IO.Path.Combine(Environment.CurrentDirectory, @"Docs\Восстановление.docx");
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
        */

        private void FrstTbGen(object sender, RoutedEventArgs e)
        {
            var filename = System.IO.Path.Combine(Environment.CurrentDirectory, @"Docs\Восстановление.docx");
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
                { "<DOG_PLATN>", FirstTab10.Text },
                { "<NEW_NAPRAVLENIE_PODGOTOVKI>", FirstTab11.Text },
                { "<NEW_NAPRAVLENIE_PROGRAMMI>", FirstTab12.Text },

            };
            paster.Process(items);
        }

        private void ScndTbGen(object sender, RoutedEventArgs e)
        {
            var filename = System.IO.Path.Combine(Environment.CurrentDirectory, @"Docs\Oтчисление по сб ж.docx");
            var paster = new WordPaster(filename);

            //Указываем ключи
            var items = new Dictionary<string, string>
            {
                { "<STUDENT_FIO>", SecondTab1.Text },
                { "<STUDENT_GROUP>", SecondTab2.Text },
                { "<STUDENT_COURSE>", SecondTab3.Text },
                { "<PAYMENT>", SecondTab4.Text },
                { "<PROGRAM>", SecondTab5.Text },
                { "<NAPRAVLEN_PROGRAM>", SecondTab6.Text },
                { "<ELEMIN_DATE>", SecondTab7.Text }

            };
            paster.Process(items);
        }
        private void ThrdTbGen(object sender, RoutedEventArgs e)
        {
            var filename = System.IO.Path.Combine(Environment.CurrentDirectory, @"Docs\Перевод с ОП на другую ОП.docx");
            var paster = new WordPaster(filename);

            //Указываем ключи
            var items = new Dictionary<string, string>
            {
                { "<STUDENT_FIO>", ThirdTab1.Text },
                { "<STUDENT_GROUP>", ThirdTab2.Text },
                { "<STUDENT_COURSE>", ThirdTab3.Text },
                { "<NAPRAVL_PODGOT>", ThirdTab4.Text },
                { "<NAPRAVL_PRORAMM>", ThirdTab5.Text },
                { "<NEW_STUDENT_COURSE>", ThirdTab6.Text },
                { "<NEW_ NAPRAVL_PODGOT>", ThirdTab7.Text },
                { "<NEW_NAPRAVL_PRORAMM>", ThirdTab8.Text },
                { "<ELEMEN_DATE>", ThirdTab9.Text },
                { "<NEW_GROUP_N_DATE>", ThirdTab10.Text },
                { "<PROTOCOL_ZASED>", ThirdTab11.Text },
                { "<PRAKTICKS>", ThirdTab12.Text }

            };
            paster.Process(items);
        }
        private void FrthTbGen(object sender, RoutedEventArgs e)
        {
            var filename = System.IO.Path.Combine(Environment.CurrentDirectory, @"Docs\приказ о зачислении в порядке перевода.docx");
            var paster = new WordPaster(filename);

            //Указываем ключи
            var items = new Dictionary<string, string>
            {
                { "<STUDENT_FIO>", FourthTab1.Text },
                { "<UNIVERSITY_FROM>", FourthTab2.Text },
                { "<FOUNDATION>", FourthTab3.Text },
                { "<STUDENT_COURSE>", FourthTab4.Text },
                { "<NAPRAVL_PODG>", FourthTab5.Text },
                { "<NAPRAVL_PROG>", FourthTab6.Text },
                { "<FORM_OF_EDUC>", FourthTab7.Text },
                { "<NEW_FOUNDATION>", FourthTab8.Text },
                { "<DATE_ELEM>", FourthTab9.Text },
                { "<STUDENT_GROUP_N_DATE>", FourthTab10.Text },
                { "<REFERENCE>", FourthTab11.Text },
                { "<PROTOCOL_ZASED>", FourthTab12.Text },
                { "<DOG_OKAZ_USLUG_OT_N>", FourthTab13.Text }

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
