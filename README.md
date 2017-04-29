# ReportBuilder

Данная программа, позволяет генерировать файл 'report.docx' на основе конфигурационного файла 'source.xml' и шаблона 'template.docx'.    
В файлах source.xml и template.docx необходимо указать: Картинки, Вложенные файлы, Таблицы, Содержание файлов
которые необходимо вставить в итоговый report.docx

\bin\Release\ - cобранная программа    
"startReportBuilder.bat" - запускает программу и генерирует "отчет-пример"

Для настройки собстенной конфигурации необходимо:

1) Настроить source.xml, указав название <тэга> и путь до файла    

![1](https://cloud.githubusercontent.com/assets/13558216/25557269/0447e83a-2d17-11e7-9f17-63f3184d10b7.png) 

2) Настроить template.docx, указав расположение <тэгов> в документе        

![2](https://cloud.githubusercontent.com/assets/13558216/25557270/05d9c524-2d17-11e7-8cb1-9800cca6c697.png)    

3) Запустить startReportBuilder.bat, программа сгенерирует 'report.docx', в котором будут добавлены Картинки, Вложенные файлы, Таблицы  

![3](https://cloud.githubusercontent.com/assets/13558216/25557271/07c38d84-2d17-11e7-94d6-e066e0663f3f.png)     

___________________
__English__

This program allows you to generate a file 'report.docx' on the basis of a configuration file 'source.xml' and template 'template.docx'.     
Files source.xml and template.docx you must specify: Pictures, File attachments, Tables, File content to be inserted in the final report.docx    

\bin\Release\ - builded program   
"sartReportBuilder.bat" - starts the program and generates an "example report"   

To make your own configuration you need:   
1) To configure source.xml name <tag> and the path to the file   
2) To configure template.docx, putting the <tag> in the document     
3) Start startReportBuilder.bat the program will generate a 'report.docx', where you will add Pictures, attachments, Tables    
