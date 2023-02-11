using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeOpenXml;
using Word = Microsoft.Office.Interop.Word;

namespace ExcelReader;

public class FileExcelReader
{
    private const string _dirName = "result_files";
    public string targetDir { get; set; }
    public string FileExcelPath { get; set; }
    public string FileWordTemplate { get; set; }

    public delegate void ExcelHandler(string message);

    public event ExcelHandler? Notify;

    public string CreateResultDirectory()
    {
        string currentPath = Directory.GetCurrentDirectory() + "\\" + _dirName;
        if (!Directory.Exists(currentPath))
        {
            Directory.CreateDirectory(currentPath);
        }
        this.targetDir = currentPath;
        return $"Create result directory {currentPath} \n";
    }

    public string CloneWordFile(string fileName)
    {
        string filePath = this.targetDir + "\\" + fileName + ".docx";
        FileInfo fileInf = new FileInfo(this.FileWordTemplate);
        if (fileInf.Exists)
        {
            fileInf.CopyTo(filePath, true);
        }
        Notify?.Invoke($"File {this.FileWordTemplate} cloned into {filePath}\n");
        return filePath;
    }

    public void WriteToWord(object[] items, string filepath)
    {
        // Получаем массив байтов из нашего файла
        byte[] textByteArray = File.ReadAllBytes(filepath);
        //Начинаем работу с потоком
        // Массив данных
        DateTime currentDate = DateTime.Now;
        string formattedDate = currentDate.ToString("dd.MM.yyyy");
        string[] arBirthday = items[12].ToString().Split(" ");
        string[] data = new string[5]
        {
            items[32].ToString(),
            items[15].ToString(),
            arBirthday[0],
            items[19].ToString() + " " + items[20].ToString(),
            formattedDate
        };
        using (MemoryStream stream = new MemoryStream())
        {
            //Записываем фаил в поток
            stream.Write(textByteArray, 0, textByteArray.Length);
            // Открываем документ из потока с возможностью редактирования
            using (WordprocessingDocument doc = WordprocessingDocument.Open(stream, true))
            {
                // Ищем все закладки в документе
                var bookMarks = FindBookmarks(doc.MainDocumentPart.Document);
                int i = 0;
                foreach (var end in bookMarks)
                {
                    // В документе встречаются какие-то служебные закладки
                    // Таким способом отфильтровываем всё ненужное
                    // end.Key содержит имена наших закладок
                    //if (end.Key != "Name" && end.Key != "Age" && end.Key != "Surname") continue;
                    // Создаём текстовый элемент
                    var textElement = new Text(data[i].ToString());
                    // Далее данный текст добавляем в закладку
                    var runElement = new Run(textElement);
                    end.Value.InsertAfterSelf(runElement);
                    i++;
                }
            }
            // Записываем всё в наш файл
            File.WriteAllBytes(filepath, stream.ToArray());
        }
    }
    
    // Получаем все закладки в документе
    // bStartWithNoEnds - словарь, который будет содержать только начало закладок,
    // чтобы потом по ним находить соответствующие им концы закладок
    // documentPart - элемент Word-документа
    // outs - конечный результат
    private Dictionary<string, BookmarkEnd> FindBookmarks(OpenXmlElement documentPart,
        Dictionary<string, BookmarkEnd> outs = null, Dictionary<string, string> bStartWithNoEnds = null)
    {
        if (outs == null)
        {
            outs = new Dictionary<string, BookmarkEnd>();
        }

        if (bStartWithNoEnds == null)
        {
            bStartWithNoEnds = new Dictionary<string, string>();
        }

        // Проходимся по всем элементам на странице Word-документа
        foreach (var docElement in documentPart.Elements())
        {
            // BookmarkStart определяет начало закладки в рамках документа
            // маркер начала связан с маркером конца закладки
            if (docElement is BookmarkStart)
            {
                var bookmarkStart = docElement as BookmarkStart;
                // Записываем id и имя закладки
                bStartWithNoEnds.Add(bookmarkStart.Id, bookmarkStart.Name);
            }

            // BookmarkEnd определяет конец закладки в рамках документа
            if (docElement is BookmarkEnd)
            {
                var bookmarkEnd = docElement as BookmarkEnd;
                foreach (var startName in bStartWithNoEnds)
                {
                    // startName.Key как раз и содержит id закладки
                    // здесь проверяем, что есть связь между началом и концом закладки
                    if (bookmarkEnd.Id == startName.Key)
                        // В конечный массив добавляем то, что нам и нужно получить
                        outs.Add(startName.Value, bookmarkEnd);
                }
            }

            // Рекурсивно вызываем данный метод, чтобы пройтись по всем элементам
            // word-документа
            FindBookmarks(docElement, outs, bStartWithNoEnds);
        }
        return outs;
    }

    public void ReadExcelFile()
    {
        using (var package = new ExcelPackage(new System.IO.FileInfo(this.FileExcelPath)))
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // Перебор всех листов в книге
            foreach (var sheet in package.Workbook.Worksheets)
            {
                Notify?.Invoke($"Start to work with {sheet.Name}\n");
                // Создание DataTable
                DataTable dataTable = new DataTable();

                // Добавление колонок в DataTable
                for (int col = 1; col <= sheet.Dimension.End.Column; col++)
                {
                    dataTable.Columns.Add(sheet.Cells[1, col].Value.ToString());
                }

                // Добавление строк в DataTable
                for (int row = 2; row <= sheet.Dimension.End.Row; row++)
                {
                    DataRow dataRow = dataTable.NewRow();
                    for (int col = 1; col <= sheet.Dimension.End.Column; col++)
                    {
                        dataRow[col - 1] = sheet.Cells[row, col].Value;
                    }
                    dataTable.Rows.Add(dataRow);
                }
                
                foreach (DataRow row in dataTable.Rows)
                {
                    var arItems = row.ItemArray.ToArray();
                    var path = this.CloneWordFile(arItems[9].ToString());
                    Notify?.Invoke($"Start write to file {arItems[9].ToString()}\n");
                    this.WriteToWord(arItems, path);
                }
            }
        }
    }
}