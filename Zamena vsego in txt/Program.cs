using System;
using System.IO;
using Aspose.Cells;
namespace Zamena_vsego_in_txt
{
    class Program
    {
        static void Main(string[] args)
        {
            char razdel = '#';
            Workbook wb = new Workbook("Test.xlsx");
            StreamWriter sw = new StreamWriter("text1.txt");
            String line = "";
            WorksheetCollection collection = wb.Worksheets;

            for(int i = 0; i < collection.Count;i++)
            {
                Worksheet worksheet = collection[i];

                int rows = worksheet.Cells.MaxDataRow;
                int col = worksheet.Cells.MaxDataColumn;

                for(int j = 0; j <= rows;j++)
                {
                    for(int k = 0; k <= col;k++)
                    {
                            if (k == 14 || k == 15)
                            {
                                if (worksheet.Cells[j, k].Value.ToString() == "" )
                                {
                                    line = line + "0" + razdel;
                                }
                                else
                                {
                                    line = line + worksheet.Cells[j, k].Value.ToString().Replace(',', '.') + razdel;
                                }
                            }
                            
                            else
                            {
                                    line = line + worksheet.Cells[j, k].Value.ToString().Replace('\n', ' ') + razdel;
                    
                            } 
                    }
                    //Console.WriteLine(line);
                    sw.WriteLine(line.Substring(0,line.Length-1));
                    line = "";
                }
            }

            sw.Close();
            Console.WriteLine("ss");
        }
    }
}
