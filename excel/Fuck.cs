using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using NPOI.OpenXml4Net.OPC;
using NPOI.XWPF.UserModel;

namespace excel
{
    public class Fuck
    {
        public static void Main1()
        {
            string filePath = "/home/kraiben/Downloads/Telegram Desktop/1.docx"; // Путь к вашему файлу

            List<string> messages = ReadWordDocument(filePath);

            List<string> emails = ExtractEmails(messages);

            // Создаем словарь для хранения адресов электронной почты по типам
            Dictionary<string, List<string>> emailDictionary = new Dictionary<string, List<string>>();

            // Ключевые слова для типов
            Dictionary<string, List<string>> keywords = new Dictionary<string, List<string>>
            {
                {"Судовладельцы", new List<string>{"Vessel", "Vessels", "MV", "Fleet", "dwt"}},
                {"Брокеры", new List<string>{"Urea", "Tn", "Fertilizer", "Fertilizers", "Steel"}}
            };

            // Инициализируем список для каждого типа
            foreach (var kvp in keywords)
            {
                emailDictionary.Add(kvp.Key, new List<string>());
            }

            // Проверяем каждый адрес на наличие ключевых слов и добавляем в соответствующий список
            foreach (string email in emails)
            {
                foreach (string message in messages)
                {
                    if (IsKeywordNearEmail(message, email, keywords))
                    {
                        foreach (var kvp in keywords)
                        {
                            foreach (string keyword in kvp.Value)
                            {
                                if (message.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                                {
                                    emailDictionary[kvp.Key].Add(email);
                                    break; // Если слово найдено, нет смысла проверять остальные типы
                                }
                            }
                        }
                        break; // Если найдено ключевое слово рядом с адресом, нет смысла искать его в остальных сообщениях
                    }
                }
            }

            // Выводим содержимое словаря
            foreach (var kvp in emailDictionary)
            {
                Console.WriteLine($"Тип: {kvp.Key}");
                Console.WriteLine("Адреса:");
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

        static List<string> ExtractEmails(List<string> messages)
        {
            List<string> emails = new List<string>();

            string pattern = @"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}";

            RegexOptions options = RegexOptions.IgnoreCase;

            foreach (string message in messages)
            {
                MatchCollection matches = Regex.Matches(message, pattern, options);

                foreach (Match match in matches)
                {
                    emails.Add(match.Value);
                }
            }

            return emails;
        }

        static bool IsKeywordNearEmail(string message, string email, Dictionary<string, List<string>> keywords)
        {
            int emailIndex = message.IndexOf(email, StringComparison.OrdinalIgnoreCase);
            if (emailIndex == -1)
                return false;

            string subString = message.Substring(Math.Max(0, emailIndex - 100), Math.Min(200, message.Length - Math.Max(0, emailIndex - 100)));
            foreach (var kvp in keywords)
            {
                foreach (string keyword in kvp.Value)
                {
                    if (subString.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }
            return false;
        }
    }
}
