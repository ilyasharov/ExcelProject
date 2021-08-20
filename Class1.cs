using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Xml;
using System.Diagnostics;
using System.Collections;
using System.Globalization;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTest
{
    static public class Class1
    {
        static public void excelWork1(Workbook excelWorkbook, _Worksheet excelWorkbookWorksheet, int LastRow)
        {
            try
            {
                HashSet<string> labelsList = new HashSet<string>(); // список Labels
                Dictionary<string, Dictionary<string, uint>> dict = new Dictionary<string, Dictionary<string, uint>>(); //словарь результатов
                
                for (int j = 2; CheckEnd(excelWorkbookWorksheet, j); j++) // перебор строк
                {
                    if (excelWorkbookWorksheet.Cells[j, 6].Value2 != null && excelWorkbookWorksheet.Cells[j, 7].Value2 != null) // проверка наличия данных в ячейках
                    {
                        string memberCell = excelWorkbookWorksheet.Cells[j, 6].Value2.ToString(); // получение текста ячеек
                        string labelCell = excelWorkbookWorksheet.Cells[j, 7].Value2.ToString();
                        if (memberCell != string.Empty && labelCell != string.Empty) // проверка наличия данных в ячейках
                        {
                            string[] members = memberCell.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries); // разделение текста
                            string[] labels = labelCell.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                            foreach (string member in members) // перебор Members
                            {
                                if (!dict.ContainsKey(member)) // не содержит member
                                {
                                    dict.Add(member, new Dictionary<string, uint>()); // добавить member в словарь
                                }
                                foreach (string label in labels) // перебор Labels
                                {
                                    if (!dict[member].ContainsKey(label)) // member не содержит label
                                    {
                                        dict[member].Add(label, 1); // // добавить label в словарь member
                                    }
                                    else
                                    {
                                        dict[member][label]++; //увеличить количество
                                    }
                                    if (!labelsList.Contains(label)) // не содержится в списке
                                    {
                                        labelsList.Add(label); // добавить в список
                                    }
                                }
                            }
                        }
                    }
                }

                excelWorkbook.Close(false); // закрыть книгу

                Console.WriteLine("Расширение вводить не нужно, файл будет сохранён в корневой каталог диска C\n");
                Console.WriteLine("Введите каталог для сохранения (только букву без двоеточия): \n");
                string name = Console.ReadLine();
                string path = name + @":\Results";

                Excel.Application excel = new Excel.Application();
                excel.SheetsInNewWorkbook = 1; // количество листов в новой книге
                excelWorkbook = excel.Workbooks.Add(); // создание книги
                excelWorkbookWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item(1); // получение листа
                excelWorkbookWorksheet.Name = "Result"; // название листа
                    
                string[] labels2 = labelsList.ToArray(); // конвертация в массив
                uint[] sum = new uint[labels2.Length]; // суммы по каждому label 
                int i = 2; // номер строки
                    
                foreach (KeyValuePair<string, Dictionary<string, uint>> item in dict) // перебор member
                    {
                    excelWorkbookWorksheet.Cells[i, 1] = item.Key; // текст member
                        for (int j = 0; j < labels2.Length; j++) // перебор label
                        {
                            if (item.Value.ContainsKey(labels2[j])) // label содержится в member
                            {
                            excelWorkbookWorksheet.Cells[i, j + 2] = item.Value[labels2[j]]; // количество
                                sum[j] += item.Value[labels2[j]]; // суммирование
                            }
                            else
                            {
                            excelWorkbookWorksheet.Cells[i, j + 2] = 0;
                            }
                        }
                        i++;
                    }
                    for (int j = 0; j < labels2.Length; j++) // перебор label
                    {
                        excelWorkbookWorksheet.Cells[1, j + 2] = labels2[j]; // текст label
                        excelWorkbookWorksheet.Cells[dict.Count + 2, j + 2] = sum[j]; // сумма по label
                    }

                    excelWorkbook.SaveAs(path); 
                    excelWorkbook.Close(false); 
                
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }

        }
        public static bool CheckEnd(_Worksheet excelWorkbookWorksheet, int row)
        {
            for (int i = 1; i <= 9; i++)
            {
                object cellValue = excelWorkbookWorksheet.Cells[row, i].Value2;
                if (cellValue != null && cellValue.ToString() != string.Empty)
                {
                    return true;
                }
            }
            return false;
        }

        //Удаление строк
        static public void excelWork2(_Worksheet excelWorkbookWorksheet, int LastRow)
        {
            for (int i = LastRow; i >= 2; i--)
            {
                if (excelWorkbookWorksheet.Cells[i, 1].Text.ToString() != @"Done 🎉")
                {
                    excelWorkbookWorksheet.Rows[i].Delete(XlDeleteShiftDirection.xlShiftUp);
                }
            }
        }

        //Удаление столбцов
        static public void excelWork3(_Worksheet excelWorkbookWorksheet, int LastColumn)
        {
            for (int i = LastColumn; i >= 1; i--)
            {
                if (excelWorkbookWorksheet.Cells[1, i].Text.ToString() == "Card URL")
                {
                    excelWorkbookWorksheet.Columns[i].Delete();
                }
                if (excelWorkbookWorksheet.Cells[1, i].Text.ToString() == "Card #")
                {
                    excelWorkbookWorksheet.Columns[i].Delete();
                }
                if (excelWorkbookWorksheet.Cells[1, i].Text.ToString() == "Points")
                {
                    excelWorkbookWorksheet.Columns[i].Delete();
                }
            }
        }

        //Выравнивание для всех ячеек
        static public void excelWork4(_Worksheet excelWorkbookWorksheet, int LastColumn, int LastRow)
        {
            for (int i = LastColumn; i >= 1; i--)
            {
                for (int j = LastRow; j >= 1; j--)
                {
                    excelWorkbookWorksheet.Cells[i, j].VerticalAlignment = XlHAlign.xlHAlignCenter;
                    excelWorkbookWorksheet.Cells[i, j].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    excelWorkbookWorksheet.Cells[i, j].Style.WrapText = true;
                }
            }
        }

        //Выравнивание для одного столбца
        static public void excelWork5(_Worksheet excelWorkbookWorksheet, int LastColumn, int LastRow)
        {
            for (int i = LastColumn; i >= 1; i--)
            {
                if (excelWorkbookWorksheet.Cells[1, i].Text.ToString() == "Description")
                {
                    for (int j = LastRow; j >= 1; j--)
                    {
                        excelWorkbookWorksheet.Rows[j].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        excelWorkbookWorksheet.Cells[i, j].VerticalAlignment = XlHAlign.xlHAlignDistributed;
                        excelWorkbookWorksheet.Cells[i, j].Style.WrapText = true;
                    }
                }
            }
        }

        //Полужирные заголовки
        static public void excelWork6(_Worksheet excelWorkbookWorksheet, int LastColumn)
        {
            for (int i = LastColumn; i >= 1; i--)
            {
                excelWorkbookWorksheet.Cells[1, i].Font.Bold = true;
            }
        }

        static public void excelWork7(_Worksheet excelWorkbookWorksheet, int LastColumn, int LastRow)
        {
            for (int i = LastRow; i >= 1; i--)
            {
                for (int j = LastColumn; j >= 1; j--) { 

                    excelWorkbookWorksheet.Cells[j, i].Style.WrapText = true;
                }
            }

        }
    }
}
