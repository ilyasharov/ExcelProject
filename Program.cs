using System;
using System.Diagnostics;
using System.Collections.Generic;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTest { 
    class Program1
    {
        static void Main()
        {
            try{

                Excel.Application excel = new Excel.Application();
                // open the concrete file:
                Excel.Workbook excelWorkbook = excel.Workbooks.Open(@"G:\2021_Май.xlsx");
                // select worksheet. NOT zero-based!!:
                Excel._Worksheet excelWorkbookWorksheet = excelWorkbook.Sheets[1];

                //Последняя строка
                int LastRow = excelWorkbook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                //Последняя колонка
                int LastColumn = excelWorkbook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

                Console.WriteLine("1 - первая часть задания\n");
                Console.WriteLine("2 - вторая часть задания\n\n\n");
                Console.WriteLine("Выберите 1 или 2");
                int selection = Convert.ToInt32(Console.ReadLine());
                switch (selection)
                {
                    case 1:

                        Console.WriteLine("Вы выбрали вариант 1");
                        Console.WriteLine("Ожидайте завершения работы программы...");

                        //Удаление строк
                        Class1.excelWork2(excelWorkbookWorksheet, LastRow);

                        //Удаление столбцов
                        Class1.excelWork3(excelWorkbookWorksheet, LastColumn);

                        //Выравнивание для всех ячеек
                        Class1.excelWork4(excelWorkbookWorksheet, LastColumn, LastRow);

                        //Выравнивание для одного столбца
                        Class1.excelWork5(excelWorkbookWorksheet, LastColumn, LastRow);

                        //Установка полужирных заголовков
                        Class1.excelWork6(excelWorkbookWorksheet, LastColumn);

                        //Выравнивание по объёму текста
                        Class1.excelWork7(excelWorkbookWorksheet, LastColumn, LastRow);

                        excelWorkbook.Save();

                        break;

                    case 2:

                        Console.WriteLine("Вы выбрали вариант 2");

                        Class1.excelWork1(excelWorkbook, excelWorkbookWorksheet, LastRow);

                        break;

                    default:
                        Console.WriteLine("Вы нажали неизвестную цифру/букву");
                        break;
                }

                excel.Quit();

                if (excel != null)
                {
                    Process[] pProcess;
                    pProcess = Process.GetProcessesByName("Excel");
                    pProcess[0].Kill();
                }

                Console.WriteLine("\n\nНажмите ENTER для выхода");
                Console.ReadLine();
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }

        }

    }
}
