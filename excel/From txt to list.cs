using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace excel;

public class From_txt_to_list
{
    public static List<string> ReadEmailsFromTxt(string filePath)
    {
        List<string> emails = new List<string>();

        using StreamReader rd = new StreamReader(filePath);
        {
            string line;
            while ((line = rd.ReadLine()) != null)
            {
                emails.Add(line);
            } 
        }
        foreach (var email in emails.Distinct())
        {
            Console.WriteLine(email);
        }

        return emails.Distinct().ToList();
    }
    

}