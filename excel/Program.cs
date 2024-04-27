using System;
using System.IO;
using OfficeOpenXml;

class Program
{
    static void Main(string[] args)
    {
        string inputFile = "input.xlsx";
        string outputFile = "output.xlsx";

        // Открываем файл Excel для чтения
        FileInfo existingFile = new FileInfo(inputFile);
        using (ExcelPackage package = new ExcelPackage(existingFile))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets["Лист1"];

            // Получаем количество строк в таблице
            const int rowCount = 276;

            if (worksheet != null)
            {
                // Проходимся по каждой строке, начиная со второй (первая строка - заголовки)
                for (int row = 2; row <= rowCount; row++)
                {
                    object cellValueObject = worksheet.Cells[row, 1].Value;
                    if (cellValueObject != null)
                    {
                        string email = cellValueObject.ToString().ToLower();

                        // Проверяем наличие ключевых слов и записываем соответствующий тип
                        if (email.Contains("chickpeas") || email.Contains("wheat") || email.Contains("coal") || email.Contains("corn") || email.Contains("cargo"))
                        {
                            worksheet.Cells[row, 2].Value = "Брокеры";
                        }
                        else if (email.Contains("handymax") || email.Contains("handy") || email.Contains("panamax") || email.Contains("supramax") || email.Contains("open"))
                        {
                            worksheet.Cells[row, 2].Value = "Судовладельцы";
                        }
                        else
                        {
                            worksheet.Cells[row, 2].Value = "Не определено";
                        }
                    }
                    else
                    {
                        // Обработка пустой ячейки
                        worksheet.Cells[row, 2].Value = "Пусто";
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
