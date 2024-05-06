using System.Text.RegularExpressions;
using NPOI.OpenXml4Net.OPC;
using NPOI.XWPF.UserModel;

namespace excel
{
    public class Fuck
    {
        public static void Main1()
        {
            //string filePath = "/home/kraiben/Downloads/Telegram Desktop/3.docx"; // Путь к вашему файлу
            string filePath = "/home/kraiben/Downloads/test.docx"; // Путь к вашему файлу

            List<string> messages = ReadWordDocument(filePath);

            // Создаем словарь для хранения адресов электронной почты по типам
            Dictionary<string, HashSet<string>> emailDictionary = new Dictionary<string, HashSet<string>>();

            // Ключевые слова для типов
            Dictionary<string, List<string>> keywords = new Dictionary<string, List<string>>
            {
                {"Vessel", new List<string>{"vessel", "mv", "fleet", "dwt", "handy", "panamax", "supramax", "open"}},
                {"Cargo", new List<string>{"urea", "tn", "fertilizer", "fertilizers", "steel", "chickpeas", "wheat", "coal", "corn", "cargo", "1000x1000", "grain"}}
            };

            // Инициализируем список для каждого типа
            foreach (var kvp in keywords)
            {
                emailDictionary.Add(kvp.Key, new HashSet<string>());
            }

            // Проверяем каждое сообщение
            foreach (string message in messages)
            {
                // Поиск адресов электронной почты в сообщении
                string pattern = @"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}";
                RegexOptions options = RegexOptions.IgnoreCase;
                MatchCollection matches = Regex.Matches(message, pattern, options);

                // Проверка каждого адреса электронной почты
                foreach (Match match in matches)
                {
                    int emailIndex = match.Index;
                    int startIndex = Math.Max(0, emailIndex - 200);
                    int endIndex = Math.Min(emailIndex + 200, message.Length - 1);
                    string subString = message.Substring(startIndex, endIndex - startIndex);
                    
                    bool emailAdded = false; // Флаг для отслеживания добавления адреса в категорию

                    // Проверяем каждый тип наличия ключевого слова в сообщении
                    foreach (var kvp in keywords)
                    {
                        foreach (string keyword in kvp.Value)
                        {
                            // Проверяем, содержится ли ключевое слово справа от адреса электронной почты
                            if (subString.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                            {
                                // Добавляем адрес электронной почты в соответствующий HashSet,
                                // если он еще не был добавлен в другую категорию
                                if (!emailAdded)
                                {
                                    emailDictionary[kvp.Key].Add(match.Value);
                                    emailAdded = true; // Устанавливаем флаг добавления адреса
                                }
                                break; // Нет смысла продолжать проверку ключевых слов для этого сообщения
                            }
                        }
                    }
                }
            }

            // Выводим содержимое словаря
            foreach (var kvp in emailDictionary)
            {
                Console.WriteLine($"Тип: {kvp.Key}");
                Console.WriteLine("Адреса электронной почты:");
                foreach (string email in kvp.Value)
                {
                    Console.WriteLine(email);
                }
                Console.WriteLine();
            }
        }

        static List<string> ReadWordDocument(string filePath)
        {
            List<string> messages = new List<string>();

            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    XWPFDocument doc = new XWPFDocument(OPCPackage.Open(fs));

                    foreach (XWPFParagraph paragraph in doc.Paragraphs)
                    {
                        string paragraphText = paragraph.Text;
                        messages.Add(paragraphText);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка при чтении документа: " + ex.Message);
            }

            return messages;
        }
    }
}
