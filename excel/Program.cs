using System;
using System.IO;
using OfficeOpenXml;

class Program
{
    static void Main()
    {
        ToExcel();
        Category();
    }
    static void ToExcel()
    {
        // Укажите путь к вашему текстовому файлу с электронными почтами
        string textFilePath = "/home/kraiben/Code/excel/excel/bin/Debug/net7.0/1.txt";

        // Укажите путь для сохранения Excel-файла
        string excelFilePath = "/home/kraiben/Code/excel/excel/bin/Debug/net7.0/input.xlsx";

        // Создаем новый пакет Excel
        ExcelPackage excelPackage = new ExcelPackage();

        // Добавляем новый лист
        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Почты");

        // Читаем почты из текстового файла и записываем их в столбец Excel-файла
        List<string> emails = new List<string>(File.ReadAllLines(textFilePath).Distinct());
        for (int i = 0; i < emails.Count; i++)
        {
            // Записываем каждую почту в ячейку в столбце A
            worksheet.Cells[i + 1, 1].Value = emails[i];
        }

        // Сохраняем Excel-файл
        excelPackage.SaveAs(new FileInfo(excelFilePath));

        Console.WriteLine("Почты успешно записаны в Excel-файл.");
    }
    static void Category()
    {
        string inputFile = "input.xlsx";
        string outputFile = "output.xlsx";

        // Открываем файл Excel для чтения
        FileInfo existingFile = new FileInfo(inputFile);
        using (ExcelPackage package = new ExcelPackage(existingFile))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets["Почты"];
            
            // Получаем количество строк в таблице
            int rowCount = worksheet.Dimension.Rows;
            
            if (worksheet != null)
            {
                // Проходимся по каждой строке, начиная со второй (первая строка - заголовки)
                for (int row = 1; row <= rowCount; row++)
                {    
                    object cellValueObject = worksheet.Cells[row, 1].Value;
                    if (cellValueObject != null)
                    {
                        string email = cellValueObject.ToString().ToLower();

                        // Проверяем наличие ключевых слов и записываем соответствующий тип
                        if (email.Contains("chickpeas") || email.Contains("wheat") || email.Contains("coal") || email.Contains("corn") || email.Contains("cargo") || email.Contains("urea") || email.Contains("tn") || email.Contains("fertlizer") || email.Contains("steel") || email.Contains("broker") || email.Contains("trade"))
                        {
                            Console.WriteLine($"{email} is belongs to брокер");
                            worksheet.Cells[row, 2].Value = "Брокеры";
                        }
                        else if (email.Contains("handymax") || email.Contains("handy") || email.Contains("panamax") || email.Contains("supramax") || email.Contains("open") || email.Contains("vessel") || email.Contains("mv") || email.Contains("vessel") || email.Contains("fleet") || email.Contains("dwt") || email.Contains("ship") || email.Contains("wave"))
                        {
                            Console.WriteLine($"{email} is belongs to судовладельцы");
                            worksheet.Cells[row, 2].Value = "Судовладельцы";
                        }
                        else
                        {
                            worksheet.Cells[row, 2].Value = "Брокеры";
                        }
                    }
                    else
                    {
                        Console.WriteLine("Nothing");
                    }
                }
                // Сохраняем изменения
                FileInfo outputFileInfo = new FileInfo(outputFile);
                package.SaveAs(outputFileInfo);
            }
        }

        Console.WriteLine("Готово!");
    }
}
