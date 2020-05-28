using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OfficeOpenXml;

namespace CSVtoXLSX
{
    class Program
    {
        static void Main(string[] args)
        {
            string[] lines;
            var list = new List<string>();
            var fileStream = new FileStream(@"C:\Users\Emre-USA\Desktop\dirlist.txt", FileMode.Open, FileAccess.Read);
            using (var streamReader = new StreamReader(fileStream, Encoding.UTF8))
            {
                string line;
                while ((line = streamReader.ReadLine()) != null)
                {
                    list.Add(line);
                }
            }
            lines = list.ToArray();

            for (int i=0;i< lines.Length;i++)
            {
                string csvFileName = lines[i];
                string excelFileName = lines[i]+".xlsx";

                string worksheetsName = "TEST";

                bool firstRowIsHeader = false;

                var format = new ExcelTextFormat();
                format.Delimiter = ',';
                format.EOL = "\r";              // DEFAULT IS "\r\n";
                                                // format.TextQualifier = '"';

                using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFileName)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(worksheetsName);
                    worksheet.Cells["A1"].LoadFromText(new FileInfo(csvFileName), format, OfficeOpenXml.Table.TableStyles.Medium27, firstRowIsHeader);
                    package.Save();
                }

                Console.WriteLine("Finished!");
                //Console.ReadLine();
            }

            
        }
    }
}
