using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;
using System.Net;

namespace ExelTEst
{
    class Program
    {
        static void Main(string[] args)
        {
            double totalgain = 0;
            double totalspend = 0;
            double totalavgnet = 0;
            double totalavgspend = 0;
            double totalnet = 0;

            var dirinfo = new DirectoryInfo((@"C:\Users\pk_bl\Documents\ExcelDocs"));
            var files = dirinfo.GetFiles("*");
            List<string> filepaths = new List<string>();
            foreach (var file in files)
            {

                filepaths.Add(file.FullName);
            }
            foreach (string file in filepaths)
            {

                Application xlApp = new Application();

                WebClient client = new WebClient();
                client.DownloadFile("getfakeexcel.com/id?=5", @"C:\Users\ryhunter\Documents\ExcelReadTest.CSV");

                Workbook xlWorkbook = xlApp.Workbooks.Open(file);//xlApp.Workbooks.Open(@"C:\Users\ryhunter\Documents\ExcelReadTest.CSV");
                _Worksheet xlWorksheet = xlWorkbook.Sheets[1] as _Worksheet;
                Range xlRange = xlWorksheet.UsedRange;

                double monthgain = 0;
                double monthspend = 0;
                double netspend = 0;

                for (int i = 2; i < xlRange.Count; i++)
                {

                    if ((xlRange.Cells[i, 4] as Range).Text.ToString() == "")
                        break;
                    double val = Convert.ToDouble((xlRange.Cells[i, 4] as Range).Text.ToString());
                    monthspend += val > 0 ? 0 : val;
                    monthgain += val > 0 ? val : 0;
                }
                netspend = monthspend + monthgain;
                DbService.SaveAggregates(file, monthgain, monthspend, netspend);

                Console.WriteLine("File Name: " + file + "\n" + "Month Spend = " + monthspend + " Month Gain = " + monthgain + " Month Net = " + netspend);
                totalspend += monthspend;
                totalgain += monthgain;
                totalnet += netspend;
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }

            Console.WriteLine("\n Average Spend = " + (totalspend / filepaths.Count) + " Average Net = " + (totalnet / filepaths.Count));
            Console.ReadKey();

        }
    }
}
