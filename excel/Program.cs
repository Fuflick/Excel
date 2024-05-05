using System.Text.RegularExpressions;
using excel;
using NPOI.XWPF.UserModel;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

class Program
{
    static void Main()
    {
        Fuck.Main1();
        /*string filePath = "/home/kraiben/Downloads//1.docx"; // Замените на путь к вашему файлу

        List<string> pages = ReadDocumentElements(filePath);

        // Выводим каждую страницу для проверки
        for (int i = 0; i < pages.Count; i++)
        {
            Console.WriteLine($"Страница {i + 1}: {pages[i]}");
            Console.WriteLine();
        }*/
    }

    static List<string> ReadDocumentElements(string filePath)
    {
        List<string> elements = new List<string>();

        using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        {
            XWPFDocument doc = new XWPFDocument(fileStream);
            string currentElement = "";

            foreach (var paragraph in doc.Paragraphs)
            {
                string text = paragraph.Text;

                // Проверяем, содержит ли текст символ перевода страницы
                if (text.Contains("\f"))
                {
                    // Добавляем текущий элемент в список
                    if (!string.IsNullOrWhiteSpace(currentElement))
                    {
                        elements.Add(currentElement.Trim());
                    }

                    // Начинаем новый элемент
                    currentElement = "";
                }
                else
                {
                    // Добавляем текст абзаца к текущему элементу
                    currentElement += text + Environment.NewLine;
                }
            }

            // Добавляем последний элемент в список
            if (!string.IsNullOrWhiteSpace(currentElement))
            {
                elements.Add(currentElement.Trim());
            }
        }

        return elements;
    }
    
}