using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Word;
using System.IO;
using Microsoft.VisualBasic.FileIO;
using System.Text;

namespace ReportBuilder.XMLparts
{

    //Класс отвещающий за замену тэгов картинками, файлами, таблицами
    public class Inserter
    {
        private String templateDOCX;
        private String resultDOCX;
        private Dictionary<string, Image> images;
        private Dictionary<string, Table> tables;
        private Dictionary<string, Attach> attaches;
        private Dictionary<string, Value> values;

        Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
        Microsoft.Office.Interop.Word.Document doc = new Microsoft.Office.Interop.Word.Document();

        public Inserter(
                        String templateDOCX,
                        String resultDOCX,
                        Dictionary<string, Image> images, 
                        Dictionary<string, Table> tables,
                        Dictionary<string, Attach> attaches,
                        Dictionary<string, Value> values
                        )
        {

            this.templateDOCX = templateDOCX;
            this.resultDOCX = resultDOCX;
            this.tables = tables;
            this.attaches = attaches;
            this.images = images;
            this.values = values;

            try
            {
                Console.WriteLine("Начинаю добавление файлов");
                Console.WriteLine(" ");

                wordApp.Visible = false;
              
                doc = wordApp.Documents.OpenNoRepairDialog(FileName: Path.GetFullPath(templateDOCX), ReadOnly: true);
                doc.Activate();

                if (this.tables != null)
                {
                    insertTables();
                }

                if (this.values != null)
                {
                    insertValues();
                }

                if (this.attaches != null){
                    insertEmbededFiles();
                }
                
                if (this.images != null)
                {
                    insertImages();
                }

              

                doc.SaveAs(Path.GetFullPath(resultDOCX));
                doc.Close();
            }            
            catch (Exception ex)
            {
                Console.WriteLine("!!! Ошибка:");
                Console.WriteLine(ex.Message);
            }
            finally
            {              
                wordApp.Quit();
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wordApp);

                Console.WriteLine("Файлы добавлены");
                Console.WriteLine();
            }


        }
        
        //Метод для добавления картинок
        public void insertImages() {

            var sel = wordApp.Selection;

            foreach (KeyValuePair<string, Image> entry in images)
            {

                string tag = "<%" + entry.Key + "%>";
                bool fileExists = File.Exists(entry.Value.path);
                bool tagExists = sel.Document.Content.Text.Contains(tag);

                if (fileExists && tagExists)
                {
                    try
                    {

                        object missing = System.Type.Missing;

                        sel.Find.Text = tag;
                        sel.Find.Replacement.Text = "";
                        sel.Range.Select();

                        object replaceAll = Microsoft.Office.Interop.Word.WdReplace.wdReplaceOne;

                        sel.Find.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;
                        sel.Find.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref replaceAll, ref missing, ref missing, ref missing, ref missing);

                        var imgPath = Path.GetFullPath(string.Format(entry.Value.path));


                        sel.InlineShapes.AddPicture(
                                FileName: imgPath,
                                LinkToFile: false,
                                SaveWithDocument: true);

                        Console.WriteLine("Добавлено: " + entry.Value.id + " " + entry.Value.path);
                        Console.WriteLine(" ");
                    }

                    catch (Exception ex)
                    {
                        Console.WriteLine("!!! Ошибка. Проверьте " + entry.Value.id + " " + entry.Value.path);
                        Console.WriteLine(ex.Message);
                    }
                }

                if (!fileExists && tagExists)
                {
                    Console.WriteLine("!!! Файл не существует: " + entry.Value.id + " " + entry.Value.path);
                    Console.WriteLine("!!! Проверьте путь до файла в source.xml");
                    Console.WriteLine(" ");
                }
                if (fileExists && !tagExists)
                {
                    Console.WriteLine("!!! Тег" + tag + " отсутствует в tempalte.docx");
                    Console.WriteLine("!!! Проверьте название тега в template.docx");
                    Console.WriteLine(" ");
                }
                if (!fileExists && !tagExists)
                {
                    Console.WriteLine("!!! Файл не существует: " + entry.Value.id + " " + entry.Value.path);
                    Console.WriteLine("!!! Проверьте путь до файла в source.xml");
                    Console.WriteLine("!!! Тег" + tag + " отсутствует в tempalte.docx");
                    Console.WriteLine("!!! Проверьте название тега в template.docx");
                    Console.WriteLine(" ");
                }


            }
        }

        //Метод для добавления вложенных файлов
        public void insertEmbededFiles()
        {

            var sel = wordApp.Selection;

            foreach (KeyValuePair<string, Attach> entry in attaches)
            {
                string tag = "<%" + entry.Key + "%>";
                bool fileExists = File.Exists(entry.Value.path);
                bool tagExists = sel.Document.Content.Text.Contains(tag);

                if (fileExists && tagExists)
                {
                    try
                    {
                        
                        object missing = System.Type.Missing;

                        sel.Find.Text = tag;
                        sel.Find.Replacement.Text = "";
                        sel.Range.Select();

                        object replaceAll = Microsoft.Office.Interop.Word.WdReplace.wdReplaceOne;

                        sel.Find.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;
                        sel.Find.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref replaceAll, ref missing, ref missing, ref missing, ref missing);

                        var attachPath = Path.GetFullPath(string.Format(entry.Value.path));
                        string fileName = Path.GetFileName(string.Format(entry.Value.path)); ;

                        if(fileName.Contains(".xlsx")|| fileName.Contains(".xls"))
                        {
                           sel.InlineShapes.AddOLEObject(
                               ClassType: "Excel.Sheet.12",
                               FileName: attachPath,
                               LinkToFile: false,
                               DisplayAsIcon: true,
                               IconFileName: Path.GetFullPath("icons/xlicons.exe"),
                               IconIndex: ref missing,
                               IconLabel: fileName,
                               Range: ref missing);
                        }
                        if (fileName.Contains(".docx") || fileName.Contains(".doc"))
                        {
                           sel.InlineShapes.AddOLEObject(
                               ClassType: ref missing,
                               FileName: attachPath,
                               LinkToFile: false,
                               DisplayAsIcon: true,
                               IconFileName: Path.GetFullPath("icons/wordicon.exe"),
                               IconIndex: ref missing,
                               IconLabel: fileName,
                               Range: ref missing);
                        }
                      if(!fileName.Contains(".docx") && !fileName.Contains(".doc") && !fileName.Contains(".xlsx") && !fileName.Contains(".xls"))
                        {
                              sel.InlineShapes.AddOLEObject(
                                  ref missing, attachPath, false, ref missing, ref missing, ref missing, ref missing, ref missing
                         );
                        }                                            

                        Console.WriteLine("Добавлено: " + entry.Value.id + " " + entry.Value.path);
                        Console.WriteLine(" ");
                    }

                    catch (Exception ex)
                    {
                        Console.WriteLine("!!! Ошибка. Проверьте: " + entry.Value.id + " " + entry.Value.path);
                        Console.WriteLine(ex.Message);
                    }
                }

                if (!fileExists && tagExists)
                {
                    Console.WriteLine("!!! Файл не существует: " + entry.Value.id + " " + entry.Value.path);
                    Console.WriteLine("!!! Проверьте путь до файла в source.xml");
                    Console.WriteLine(" ");
                }
                if (fileExists && !tagExists)
                {
                    Console.WriteLine("!!! Тег" + tag + " отсутствует в tempalte.docx");
                    Console.WriteLine("!!! Проверьте название тега в template.docx");
                    Console.WriteLine(" ");
                }
                if (!fileExists && !tagExists)
                {
                    Console.WriteLine("!!! Файл не существует: " + entry.Value.id + " " + entry.Value.path);
                    Console.WriteLine("!!! Проверьте путь до файла в source.xml");
                    Console.WriteLine("!!! Тег" + tag + " отсутствует в tempalte.docx");
                    Console.WriteLine("!!! Проверьте название тега в template.docx");
                    Console.WriteLine(" ");
                }

            }

        }

        //Метод для добавления таблицы
        public void insertTables() {

            var sel = wordApp.Selection;

            foreach (KeyValuePair<string, Table> entry in tables)
            {
                string tag = "<%" + entry.Key + "%>";
                bool fileExists = File.Exists(entry.Value.path);
                bool tagExists = sel.Document.Content.Text.Contains(tag);

                if (fileExists && tagExists)
                {
                    try
                    {
                        List<string[]> csv = readCsv(entry.Value.path);
                       

                        int rowsCount = csv.Count();
                        int columnsCount = csv.ElementAt(0).Length;

                        object missing = System.Type.Missing;
                    
                        sel.Find.Text = tag;
                        sel.Find.Replacement.Text = "";
                        sel.Range.Select();
                        object replaceAll2 = Microsoft.Office.Interop.Word.WdReplace.wdReplaceOne;

                        sel.Find.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref replaceAll2, ref missing, ref missing, ref missing, ref missing);

                        Microsoft.Office.Interop.Word.Table objTable = buildTable(csv, rowsCount, columnsCount);

                        Console.WriteLine("Добавлено: " + entry.Value.id + " " + entry.Value.path);
                        Console.WriteLine(" ");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("!!! Ошибка. Проверьте: " + entry.Value.id + " " + entry.Value.path);
                        Console.WriteLine(ex.Message);
                    }
                }

                if (!fileExists && tagExists)
                {
                    Console.WriteLine("!!! Файл не существует: " + entry.Value.id + " " + entry.Value.path);
                    Console.WriteLine("!!! Проверьте путь до файла в source.xml");
                    Console.WriteLine(" ");
                }
                if (fileExists && !tagExists)
                {
                    Console.WriteLine("!!! Тег " + tag + " отсутствует в tempalte.docx");
                    Console.WriteLine("!!! Проверьте название тега в template.docx");
                    Console.WriteLine(" ");
                }
                if (!fileExists && !tagExists)
                {
                    Console.WriteLine("!!! Файл не существует: " + entry.Value.id + " " + entry.Value.path);
                    Console.WriteLine("!!! Проверьте путь до файла в source.xml");
                    Console.WriteLine("!!! Тег " + tag + " отсутствует в tempalte.docx");
                    Console.WriteLine("!!! Проверьте название тега в template.docx");
                    Console.WriteLine(" ");
                }

            }


        }

        public void insertValues()
        {
            var sel = wordApp.Selection;

            foreach (KeyValuePair<string, Value> entry in values)
            {
        
                string tag = "<%" + entry.Key + "%>";
                bool fileExists = File.Exists(entry.Value.path);
                bool tagExists = sel.Document.Content.Text.Contains(tag);
                


                if (fileExists && tagExists)
                {
                    try
                    {
                        if (entry.Value.path.Contains(".docx") || entry.Value.path.Contains(".doc") ||
                           entry.Value.path.Contains(".xlsx") || entry.Value.path.Contains(".xls"))
                        {

                            object missing = System.Type.Missing;

                            sel.Find.Text = tag;
                            sel.Find.Replacement.Text = "";
                            sel.Range.Select();

                            object replaceAll = Microsoft.Office.Interop.Word.WdReplace.wdReplaceOne;

                            sel.Find.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;
                            sel.Find.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing, ref missing, ref missing,
                                ref replaceAll, ref missing, ref missing, ref missing, ref missing);
                   
                            var attachPath = Path.GetFullPath(string.Format(entry.Value.path));

                            sel.InlineShapes.AddOLEObject(
                              ClassType: ref missing,
                              FileName: attachPath,
                              LinkToFile: false,
                              DisplayAsIcon: ref missing,
                              IconFileName: ref missing,
                              IconIndex: ref missing,
                              IconLabel: ref missing,
                              Range: ref missing);
                            
                        }
                        else {

                            List<string> b = readFile(entry.Value.path);
                            string combindedString = string.Join("\x0B", b.ToArray());

                            object missing = System.Type.Missing;

                            sel.Find.Text = tag;
                            sel.Find.Replacement.Text = combindedString;
                            sel.Range.Select();

                            object replaceAll = Microsoft.Office.Interop.Word.WdReplace.wdReplaceOne;

                            sel.Find.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;
                            sel.Find.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing, ref missing, ref missing,
                                ref replaceAll, ref missing, ref missing, ref missing, ref missing);

                            Console.WriteLine("Добавлено: " + entry.Value.id + " " + entry.Value.path);
                            Console.WriteLine(" ");
                            
                        }
                        

                    }

                    catch (Exception ex)
                    {
                        Console.WriteLine("!!! Ошибка. Проверьте " + entry.Value.id + " " + entry.Value.path);
                        Console.WriteLine(ex.Message);
                    }
                }

                if (!fileExists && tagExists)
                {
                    Console.WriteLine("!!! Файл не существует: " + entry.Value.id + " " + entry.Value.path);
                    Console.WriteLine("!!! Проверьте путь до файла в source.xml");
                    Console.WriteLine(" ");
                }
                if (fileExists && !tagExists)
                {
                    Console.WriteLine("!!! Тег" + tag + " отсутствует в tempalte.docx");
                    Console.WriteLine("!!! Проверьте название тега в template.docx");
                    Console.WriteLine(" ");
                }
                if (!fileExists && !tagExists)
                {
                    Console.WriteLine("!!! Файл не существует: " + entry.Value.id + " " + entry.Value.path);
                    Console.WriteLine("!!! Проверьте путь до файла в source.xml");
                    Console.WriteLine("!!! Тег" + tag + " отсутствует в tempalte.docx");
                    Console.WriteLine("!!! Проверьте название тега в template.docx");
                    Console.WriteLine(" ");
                }


            }
        }

        //Метод для чтения .csv
        public List<string[]> readCsv(string path)
        {
            
            try
            {
                List<string[]> csv = new List<string[]>();

                using (TextFieldParser parser = new TextFieldParser(path, Encoding.GetEncoding("windows-1251")))
                {
                    parser.TextFieldType = FieldType.Delimited;
                    parser.SetDelimiters(";");

                    while (!parser.EndOfData)
                    {

                        //Process row
                        string[] fields = parser.ReadFields();

                        csv.Add(fields);                                             

                    }
                    
                }
             
                return csv;
            }
            catch (Exception ex)
            {
                Console.WriteLine("!!! Ошибка. При чтении таблицы из файла: " +   path);
                Console.WriteLine(ex.Message);

                List<string[]> csv = null;
                return csv;

            }
        }

        //Метод для создания таблицы
        public Microsoft.Office.Interop.Word.Table buildTable(List<string[]> csv, int rowsCount, int columnsCount)
        {
            try
            {

            Microsoft.Office.Interop.Word.Table objTable;
            object oMissing = System.Reflection.Missing.Value;
            var sel = wordApp.Selection;

            Microsoft.Office.Interop.Word.Range wrdRng = sel.Range;
            Object defaultTableBehavior = Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior;
            Object autoFitBehavior = Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow;

            objTable = doc.Tables.Add(wrdRng, rowsCount, columnsCount, defaultTableBehavior, autoFitBehavior);
            objTable.Range.ParagraphFormat.SpaceAfter = 7;
                
            int i = 0;
            int j = 0;
                
            string strText;
            for (i = 0; i <= rowsCount - 1; i++)
                for (j = 0; j <= columnsCount - 1; j++)
                {
                    strText = csv.ElementAt(i).ElementAt(j);

                    objTable.Cell(i + 1, j + 1).Range.Text = strText;
    
                    }
           // objTable.Rows[1].Range.Font.Bold = 1;
          //  objTable.Rows[1].Range.Font.Italic = 1;

            objTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            objTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;

            return objTable;

        }
            catch (Exception ex)
            {                            
                Console.WriteLine("!!! Ошибка. При создании таблицы");
                Console.WriteLine(ex.Message);

                Microsoft.Office.Interop.Word.Table objTable = null;
                return objTable;
            }


        }


        public List<string> readFile(string path)
        {
            List<string> textFile = new List<string>();
            try
            {                                           
                var fileStream = new FileStream(Path.GetFullPath(path), FileMode.Open, FileAccess.Read);
                using (var streamReader = new StreamReader(fileStream, Encoding.GetEncoding("windows-1251")))
                {
                    string line;
                    while ((line = streamReader.ReadLine()) != null)
                    {
                       textFile.Add(line);
                    }

                }

                return textFile;
            }
            catch (Exception ex)
            {
                Console.WriteLine("!!! Ошибка. При чтении файла: " + path);
                Console.WriteLine(ex.Message);

                textFile = null;
                return textFile;

            }
        }

    }
}

