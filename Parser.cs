using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;


namespace ProizvodkaWPF
{
    internal class WordPaster
    {
        private FileInfo _fileInfo;

        //Чекаем есть ли шаблон
        public WordPaster(string fileName)
        {
            if (File.Exists(fileName))
            {
                _fileInfo = new FileInfo(fileName);
            }
            else
            {
                throw new ArgumentException("File not exist");
            }
        }

        //Даем список ключей
        internal bool Process(Dictionary<string, string> items)
        {
            

            Word.Application app = null;

            try
            {
                app = new Word.Application();
                Object file = _fileInfo.FullName;
                Object missing = Type.Missing;
                app.Documents.Open(file);

                //ищем и меняем ключи
                foreach (var item in items)
                {
                    Word.Find find = app.Selection.Find;
                    find.Text = item.Key;
                    find.Replacement.Text = item.Value;

                    Object warp = Word.WdFindWrap.wdFindContinue;
                    Object replace = Word.WdReplace.wdReplaceAll;

                    find.Execute(FindText: Type.Missing,
                        MatchCase: false,
                        MatchWholeWord: false,
                        MatchWildcards: false,
                        MatchSoundsLike: missing,
                        MatchAllWordForms: false,
                        Forward: true,
                        Format: false,
                        ReplaceWith: missing, Replace: replace);
                }

                //Сохраняемся
                var BufferFileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), DateTime.Now.ToString("yyyy.MM.dd HH;mm;ss ") + _fileInfo.Name);
                Object newFileName = BufferFileName;
                app.ActiveDocument.SaveAs2(newFileName);
                app.ActiveDocument.Close();

                System.Windows.MessageBox.Show($"Документ сохранён! \n Путь к файлу: {BufferFileName}");

                return true;
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            finally
            {
                if (app != null)
                {
                    app.Quit();
                }
            }
            return false;
        }
    }
}
