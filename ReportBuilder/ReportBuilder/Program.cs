using ReportBuilder.Controller;
using ReportBuilder.XMLparts;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace ReportBuilder
//Author: Steshenko MA https://github.com/SteshenkoMA 
{
    class Program
    {
        private static Dictionary<string, Image> images = new Dictionary<string, Image>();
        private static Dictionary<string, Table> tables = new Dictionary<string, Table>();
        private static Dictionary<string, Attach> attaches = new Dictionary<string, Attach>();

        [STAThread]
        static void Main(string[] args)
        {
            DateTime appStart = DateTime.Now;
            Console.WriteLine();
            Console.Out.WriteLine("Начинаю работу " + appStart);
            Console.WriteLine();
            Console.WriteLine("Загружаю аргументы: template.docx, sourses.xml, result.docx");
            Console.WriteLine();

            if (args == null)
            {
                Console.WriteLine("Укажите аргументы: template.docx, sourses.xml, result.docx");
                Console.WriteLine("Аргументов должно быть ровно 3");
                Console.ReadLine();
            }
            if(args.Length == 3)
            {
               
                String soursesXML = args[0];
                String templateDOCX = args[1];
                String resultDOCX = args[2];

                

                bool soursesXML_Exists = File.Exists(soursesXML);
                bool templateDOCX_Exists = File.Exists(templateDOCX);
                create_reportDOCX(resultDOCX);
                bool resultDOCX_Exists = File.Exists(resultDOCX);

                if (soursesXML_Exists && templateDOCX_Exists && resultDOCX_Exists)
                {
                    Console.WriteLine("sourses.xml - " + soursesXML);
                    Console.WriteLine("template.docx - " + templateDOCX);
                    Console.WriteLine("result.docx - " + resultDOCX);
                    Console.WriteLine();

                   
                    ConfigAnalyser config = new ConfigAnalyser(soursesXML);

                    Inserter inserter = new Inserter(templateDOCX,
                                                              resultDOCX,
                                                              config.get_images(),
                                                              config.get_tables(),
                                                              config.get_attaches(),
                                                              config.get_values()
                                                              );
                    

                }
            }
            else {
                Console.WriteLine("Проверьте аргументы: template.docx, sourses.xml, result.docx");
                Console.WriteLine("Аргументов должно быть ровно 3");
                Console.WriteLine("");
           }

            DateTime appEnd = DateTime.Now;
                        
            Console.Out.WriteLine("Работа программы завершена " + appEnd + " " + (appEnd - appStart));
            Console.Out.WriteLine("Для выхода нажмите Enter");
            Console.In.ReadLine();
        }

        private static void create_reportDOCX(String resultDOCX) {

            Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();
            try
            {
                winword.Visible = false;
                object missing = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Word.Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);
                document.SaveAs(Path.GetFullPath(resultDOCX));
                document.Close();
            }
            catch {
                Console.Out.WriteLine("!!! Ошибка при создании: " + resultDOCX);
                Console.Out.WriteLine("!!! Проверьте, что папки для создания файла существуют и к ним есть доступ");
                Console.Out.WriteLine();
            }
            finally
            {
                winword.Quit();
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(winword);
            }
        }

    }
}


