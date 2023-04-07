using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.Text;

namespace HelpMe.Model
{
    //Валюты
    internal class Currencies
    {
        public static IDictionary<int, string> All = new Dictionary<int, string>();

        public static bool CurrenciesInitFromCSV(string pathToCSV)
        {
            All.Clear();
            TextFieldParser parser = new TextFieldParser(pathToCSV, Encoding.GetEncoding("windows-1251"));
            parser.TextFieldType = FieldType.Delimited;
            parser.SetDelimiters(",");
            parser.ReadFields();
            while (!parser.EndOfData)
            {
                string[] fields = parser.ReadFields();
                foreach (string field in fields)
                {
                    string[] theHolyBebra = field.Split(';');
                    All.Add(Int32.Parse(theHolyBebra[0]), theHolyBebra[1]);
                }
            }
            if (All.Count > 0)
            {
                return true;
            }
            return false;
        }
    }
}
