using ReportBuilder.XMLparts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace ReportBuilder.Controller
{
    //Класс отвещающий за парсинг sourses.xml
    public class ConfigAnalyser
    {
        private Dictionary<string, Image> images = new Dictionary<string, Image>();
        private Dictionary<string, Table> tables = new Dictionary<string, Table>();
        private Dictionary<string, Attach> attaches = new Dictionary<string, Attach>();
        private Dictionary<string,Value> values = new Dictionary<string, Value>();
        private String soursesXMLPath;
        private StreamReader sr;
               
        XmlSerializer serializer = new XmlSerializer(typeof(Sourses));

        public ConfigAnalyser(String soursesXMLPath) {

            this.soursesXMLPath = soursesXMLPath;
            sr = new StreamReader(this.soursesXMLPath, System.Text.Encoding.GetEncoding("Windows-1251"), true);
            
            analyseConfig();
        }
        
        public void analyseConfig()
        {
            try
            {
                using (TextReader reader = sr)
                {
                    Sourses result = (Sourses)serializer.Deserialize(reader);
                    set_images(result);
                    set_tables(result);
                    set_attaches(result);
                    set_values(result);

                    Console.WriteLine("Конфигурационный файл sourses.xml успешно обработан");
                    Console.WriteLine(" ");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("!!! Ошибка при чтении: sourses.xml");
                Console.WriteLine(ex.Message);
            }

        }

        private void set_images(Sourses result)
        {
            try
            {
                List<Image> imgs = result.images.data;

                foreach (Image n in imgs)
                {
                    images.Add(n.id, n);
                }
            }
            catch {
                images = null;
            }

        }

        public Dictionary <string, Image> get_images() {
            return images;
        }

        private void set_tables(Sourses result)
        {
            try
            {
                List<Table> tbls = result.tables.data;

                foreach (Table n in tbls)
                {
                    tables.Add(n.id, n);
                }
            }
            catch {
                tables = null;
            }

        }

        public Dictionary<string, Table> get_tables()
        {
            return tables;
        }

        private void set_attaches(Sourses result)
        {
            try
            {
                List<Attach> attchs = result.attaches.data;

                foreach (Attach n in attchs)
                {
                    attaches.Add(n.id, n);
                }
            }
            catch {
                attaches = null;
            }

        }

        public Dictionary<string, Attach> get_attaches()
        {
            return attaches;
        }

        private void set_values(Sourses result)
        {
            try
            {
                List<Value> vals = result.values.data;

                foreach (Value n in vals)
                {
                    values.Add(n.id, n);
                }
            }
            catch
            {
                values = null;
            }
        }
        public Dictionary<string, Value> get_values()
        {
            return values;
        }

    }
}
