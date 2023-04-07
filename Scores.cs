using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace HelpMe.Model
{
    //Счета
    internal class Scores
    {
        public static List<string> All = new List<string>();

        public static string GetLastScoresFilePath()
        {
            string pathToScoreFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Счета";
            if (Directory.Exists(pathToScoreFolder))
            {
                FileInfo[] files = new DirectoryInfo(pathToScoreFolder).GetFiles("*Счёт*");
                FileInfo resultFile = files.OrderByDescending(file => file.LastWriteTime).ToArray()[0];
                if (Properties.Settings.Default.DateTimeOfLastFile.ToShortDateString() == resultFile.LastWriteTime.ToShortDateString())
                {
                    return null;
                }
                Properties.Settings.Default.DateTimeOfLastFile = resultFile.LastWriteTime;
                return resultFile.FullName;
            } else {
                MessageBox.Show("На рабочем столе нету папки Счета");
                return null;
            }  
        }

        public static bool ScoresInitFromTXT(string pathToTXT)
        {
            All.Clear();
            StreamReader reader = File.OpenText(pathToTXT);
            string line;
            while ((line = reader.ReadLine()) != null)
            {
                All.Add(line);
            }
            if (All.Count > 0)
            {
                return true;
            }
            return false;
        }
    }
}
