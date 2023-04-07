using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace HelpMe
{
    //Ммм дата
    internal class Datas
    {
        public class Data
        {
            public string Name;
            public string EtoBase;
            public string Score;
            public string TypeOfOperation;
            public int CodeOfCurr;
            public double Summ;
            public int? CodeOfOperation;

            public Data() { }

            public Data(string name, string etoBase, string score, string typeOfOperation, int codeOfCurr, double summ, int? codeOfOperation)
            {
                Name = name;
                EtoBase = etoBase;
                Score = score;
                TypeOfOperation = typeOfOperation;
                CodeOfCurr = codeOfCurr;
                Summ = summ;
                CodeOfOperation = codeOfOperation;
            }

            public override string ToString()
            {
                return
                    "Наименование: " + Name +
                    "\nБаза: " + EtoBase +
                    "\nСчёт: " + Score +
                    "\nТип операции: " + TypeOfOperation +
                    "\nКод валюты: " + CodeOfCurr +
                    "\nСумма: " + Summ +
                    "\nКод операции: " + CodeOfOperation + "\n";
            }
        }

        public static List<Data> All;

        public static string Text(Excel.Range range, int row, int column)
        {
            return (range.Cells[row, column] == null || range.Cells[row, column].Value2 == null) ? null : (range.Cells[row, column] as Excel.Range).Value2.ToString();
        }


        public static bool DatasInitFromExcel(string pathToExcel)
        {
            object rOnly = true;
            object SaveChanges = false;
            object MissingObj = System.Reflection.Missing.Value;

            Excel.Application app = new Excel.Application();
            Excel.Workbooks workbooks = null;
            Excel.Workbook workbook = null;
            Excel.Sheets sheets = null;

            List<Data> list = new List<Data>();

            try
            {
                workbook = app.Workbooks.Open(pathToExcel);
                Worksheet worksheet = (Worksheet)workbook.Worksheets[3]; //Это тут беру Выписки
                Excel.Range UsedRange = worksheet.UsedRange; // Получаем диапазон используемых на странице ячеек
                Excel.Range urRows = UsedRange.Rows; // Получаем строки в используемом диапазоне
                Excel.Range urColums = UsedRange.Columns; // Получаем столбцы в используемом диапазоне
                string name;
                string lastName = "";
                int? code = null;
                for (int i = 3; i <= urRows.Count; i++)
                {
                    if(Text(UsedRange, i, 2)!= null)
                    {
                        if(Text(UsedRange, i, 1) != null)
                        {
                            name = Text(UsedRange, i, 1);
                            lastName = Text(UsedRange, i, 1);
                        } else
                        {
                            name = lastName;
                        }

                        if (Text(UsedRange, i, 7) != null)
                        {
                            code = Int32.Parse(Text(UsedRange, i, 7));
                        }
                        else
                        {
                            code = null;
                        }
                        list.Add(new Data(
                            name,
                            Text(UsedRange, i, 2),
                            Text(UsedRange, i, 3),
                            Text(UsedRange, i, 4),
                            Int32.Parse(Text(UsedRange, i, 5)),
                            Double.Parse(Text(UsedRange, i, 6)),
                            code
                            ));

                    }
                }
                // Очистка неуправляемых ресурсов на каждой итерации
                if (urRows != null) Marshal.ReleaseComObject(urRows);
                if (urColums != null) Marshal.ReleaseComObject(urColums);
                if (UsedRange != null) Marshal.ReleaseComObject(UsedRange);
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Возникло исключение: " + ex.Message);
            }
            finally
            {
                if (sheets != null) Marshal.ReleaseComObject(sheets);
                if (workbook != null)
                {
                    workbook.Close(SaveChanges);
                    Marshal.ReleaseComObject(workbook);
                    workbook = null;
                }
                if (workbooks != null)
                {
                    workbooks.Close();
                    Marshal.ReleaseComObject(workbooks);
                    workbooks = null;
                }
                if (app != null)
                {
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                    app = null;
                }
            }
            All = list;
            if(All.Count > 0)
            {
                return true;
            }
            return false;
        }
    }
}
