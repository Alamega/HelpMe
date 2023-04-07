using System.Collections.Generic;
using System.IO;

namespace HelpMe
{
    //Адреса
    internal class Addresses
    {
        public static IDictionary<string, string> All = new Dictionary<string, string>();

        public static bool AddressesInitFromTXT(string pathToTXT)
        {
            All.Clear();
            StreamReader reader = File.OpenText(pathToTXT);
            string line;
            reader.ReadLine();
            while ((line = reader.ReadLine()) != null)
            {
                string[] theHolyBebra = line.Split(':');
                All.Add(theHolyBebra[0], theHolyBebra[1]);
            }
            if (All.Count > 0)
            {
                return true;
            }
            return false; 
        }
    }
}
