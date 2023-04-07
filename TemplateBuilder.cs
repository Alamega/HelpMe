using HelpMe.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xceed.Document.NET;
using Xceed.Words.NET;
using static HelpMe.Datas;

namespace HelpMe
{
    internal class TemplateBuilder
    {
        public static void Replace(DocX doc,string oldValue, string newValue) {
            StringReplaceTextOptions replaceTextOptions = new StringReplaceTextOptions();
            replaceTextOptions.SearchValue = oldValue;
            replaceTextOptions.NewValue = newValue;
            doc.ReplaceText(replaceTextOptions);
        }

        public static void BuildScore(string score)
        {
            System.Collections.Generic.IEnumerable<Data> dataFromScore = Datas.All.Where(d => d.Score == score);

            if (dataFromScore.Count() > 0) {
                string fileName = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Счета").GetFiles("*Счёт*").OrderByDescending(file => file.LastWriteTime).ToArray()[0].LastWriteTime.ToShortDateString() + "_Cчёт" + score + "_готов.docx";
                File.Copy(Directory.GetCurrentDirectory() + "\\template.docx", Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Готовые счета\\" + fileName, true);
                DocX doc = DocX.Load(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Готовые счета\\" + fileName);

                Replace(doc, "{name}", dataFromScore.ElementAt(0).Name); //Имя клиента 
                Replace(doc, "{postal_address}", Addresses.All[dataFromScore.ElementAt(0).EtoBase]); //Почтовый адрес 
                Replace(doc, "{bank_account}", score); //Счёт 

                IDictionary<string, double> currAndSummaOfAdd = new Dictionary<string, double>(); //Сумма пополнений
                IDictionary<string, double> currAndSummaOfDec = new Dictionary<string, double>(); //Сумма пополнений
                IDictionary<string, double> currAndSummaOfStopped = new Dictionary<string, double>(); //Сумма пополнений

                int iM = 0;
                bool block = false;
                int countOfBlocked = 0;

                for (int i = 0; i < dataFromScore.Count(); i++)
                {
                    if(dataFromScore.ElementAt(i).TypeOfOperation != "Блокировка" && dataFromScore.ElementAt(i).CodeOfOperation != 52)
                    {
                        if (i + 1 != dataFromScore.Count()) { doc.Tables[0].InsertRow(i + 1 - iM); };
                        doc.Tables[0].Rows[i + 1 - iM].Cells[0].Paragraphs[0].InsertText(dataFromScore.ElementAt(i).Score);
                        doc.Tables[0].Rows[i + 1 - iM].Cells[1].Paragraphs[0].InsertText(dataFromScore.ElementAt(i).TypeOfOperation);
                        doc.Tables[0].Rows[i + 1 - iM].Cells[2].Paragraphs[0].InsertText(dataFromScore.ElementAt(i).Summ + " " + Currencies.All[dataFromScore.ElementAt(i).CodeOfCurr]);
                    } else {
                        block = true;
                        iM++;
                    }

                    string curr = Currencies.All[dataFromScore.ElementAt(i).CodeOfCurr];

                    if(dataFromScore.ElementAt(i).CodeOfOperation == 52 || dataFromScore.ElementAt(i).TypeOfOperation == "Блокировка")
                    {
                        if (currAndSummaOfStopped.ContainsKey(curr)) { 
                            currAndSummaOfStopped[curr] = currAndSummaOfStopped[curr] + dataFromScore.ElementAt(i).Summ; 
                        }
                        else
                        { 
                            currAndSummaOfStopped.Add(curr, dataFromScore.ElementAt(i).Summ); 
                        }
                        countOfBlocked++;
                    } else
                    {
                        switch (dataFromScore.ElementAt(i).TypeOfOperation)
                        {
                            case "Поступление":
                                if (currAndSummaOfAdd.ContainsKey(curr)) { 
                                    currAndSummaOfAdd[curr] = currAndSummaOfAdd[curr] + dataFromScore.ElementAt(i).Summ; 
                                }
                                else { 
                                    currAndSummaOfAdd.Add(curr, dataFromScore.ElementAt(i).Summ); 
                                }
                                break;
                            case "Списание":
                                if (currAndSummaOfDec.ContainsKey(curr)) { 
                                    currAndSummaOfDec[curr] = currAndSummaOfDec[curr] + dataFromScore.ElementAt(i).Summ; 
                                }
                                else { 
                                    currAndSummaOfDec.Add(curr, dataFromScore.ElementAt(i).Summ);
                                }
                                break;
                            default:
                                break;
                        }
                    }             
                }

                string summaAdd = "";
                string summaDec = "";
                string summaStopped = "";
                foreach (var curr in currAndSummaOfAdd)
                {
                    summaAdd += curr.Key + " " + curr.Value + " ";
                }

                foreach (var curr in currAndSummaOfDec)
                {
                    summaDec += curr.Key + " " + curr.Value + " ";
                }

                foreach (var curr in currAndSummaOfStopped)
                {
                    summaStopped += curr.Key + " " + curr.Value + " ";
                }

                if(summaAdd == "") { summaAdd = "отсутствует"; }
                if(summaDec == "") { summaDec = "отсутствует"; }

                Replace(doc, "{deposits}", summaAdd); //Пополнений на сумму 
                Replace(doc, "{withdrawal}", summaDec); //Выводов на сумму 
                if(block) {  Replace(doc, "{blocked}", "Количество заблокированных операций по вашему счёту " + score + " равно " + countOfBlocked + " на общую сумму: " + summaStopped); } else { Replace(doc, "{blocked}", ""); } 
                Replace(doc, "{date_time}", DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString()); //Дата и время 

                doc.Save();

                MyLogger.Instance.Info("Собран файл: " + Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Готовые счета\\" + fileName);
            }
        }
    }
}
