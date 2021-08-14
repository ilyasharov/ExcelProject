using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Diagnostics;
using System.Collections;
using System.Globalization;
using System.Collections.Generic;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.XSSF.UserModel;
using NPOI.XSSF.Model;
using NPOI.Util;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.OpenXml4Net.OPC;
using NPOI.SS.Util;
using NPOI.SS.Formula;
using NPOI.XSSF.UserModel.Helpers;
using NPOI.SS.Formula.UDF;
using NPOI.OpenXmlFormats;
using NPOI.OpenXml4Net.Exceptions;
using NPOI.SS;

namespace ExcelTest { 
    class Program1
    {
        static void Main()
        {
            try{

                Excel.Application excel = new Excel.Application();
                // open the concrete file:
                Excel.Workbook excelWorkbook = excel.Workbooks.Open(@"G:\test.xlsx");
                // select worksheet. NOT zero-based!!:
                Excel._Worksheet excelWorkbookWorksheet = excelWorkbook.Sheets[1];
                /********NPOI********/

                IWorkbook workbook;
                string fileName = @"G:\test.xlsx";
                XSSFWorkbook xssfwb;
                using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    xssfwb = new XSSFWorkbook(fs);
                    if (fileName.IndexOf(".xlsx") > 0)
                        workbook = new XSSFWorkbook(fs);
                    else if (fileName.IndexOf(".xls") > 0)
                        workbook = new HSSFWorkbook(fs);
                }

                var sheet = workbook.GetSheetAt(0);

                int LastRow = excelWorkbook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                int LastColumn = excelWorkbook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

                //var excelcells = excelWorkbookWorksheet.get_Range("A1", Type.Missing);
                //var sStr = Convert.ToString(excelcells.Value2);
                //string Text = sStr + " ";
                //Console.WriteLine($"{Text}\n");

                //Excel.Range range1 = excelWorkbookWorksheet.Range[excelWorkbookWorksheet.Cells[1, 1], excelWorkbookWorksheet.Cells[99, 99]];

                //List<Excel.Range> listOfCells = excelWorkbookWorksheet.Cells.Cast<Excel.Range>().ToList<Excel.Range>();

                //int[] str = new string[100];

                /*
                for( int i = 0; i < 20; ++i)
                {
                    str[i] = "one" + i;
                }

                for (int j = 0; j < 20; ++j)
                {
                    Console.WriteLine(str[j] + "\n");
                }
                */
                var temp = "One";
                for (int c = LastColumn; c >= 1; c--)
                {
                    if (excelWorkbookWorksheet.Columns[c].Text.ToString() == "c") {

                        //Получение одной ячейки как ранга
                        Excel.Range forYach = excelWorkbookWorksheet.Rows[c] as Excel.Range;

                        //Получаем значение из ячейки и преобразуем в строку
                        string yach = Convert.ToString(forYach.Value2);
                        //temp = yach;
                        Console.WriteLine("String: {0}\n", yach);
                    }

                    
                }

                /*
                
                for (int i = 1; i <= list.Length; i++)
                {
                    Console.WriteLine(list[i]);
                }
                */

                /**************************1 - Удаление строк******************************/
                /*
                for (int i = LastRow; i >= 1; i--)
                {
                    if (excelWorkbookWorksheet.Cells[i, 1].Text.ToString() != "Done")
                    {
                        excelWorkbookWorksheet.Rows[i].Delete();
                    }
                }
                */
                /********************************************************/

                /**************************2 - Удаление столбцов******************************/
                /*
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
                */
                /************************************************************/

                /**********************4 - Выравнивание для всех ячеек**********************/
                /*
                for (int i = LastColumn; i >= 1; i--)
                {
                     for (int j = LastRow; j >= 1; j--)
                     {
                        excelWorkbookWorksheet.Cells[i, j].VerticalAlignment = XlHAlign.xlHAlignCenter;
                        excelWorkbookWorksheet.Cells[i, j].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        excelWorkbookWorksheet.Cells[i, j].Style.WrapText = true;                      
                     }
                }
                */
                /***********************3 - Выравнивание для одного столбца*********************/
                /*
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
                */
                /************************************************************/

                /************************5 - Полужирные заголовки**************************/
                /*
                for (int i = LastColumn; i >= 1; i--)
                {
                    excelWorkbookWorksheet.Cells[1, i].Font.Bold = true;
                }
                */
                /************************************************************/

                // save changes (!!):
                excelWorkbook.Save();

                // cleanup:
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
