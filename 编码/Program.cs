using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;

using System.Collections;
using System.Text.RegularExpressions;

namespace 编码
{
    class Program
    {

        public static string path = @"c:\users\1\documents\visual studio 2015\Projects\编码\编码\清单.xls";
        static void Main(string[] args)
        {
            //GetDatas(path);
            if (File.Exists(path))
            {
                var lst= ExcelHelper.operation.ExcelToLis(path, "序号", "合计");
               // using ( FileStream FS = new FileStream(@"商品编码.txt", FileMode.Open))
             
              
                ExcelHelper.operation.CreatFile(lst);
            } 
          
          

        }

        //private static void GetDatas(string path)
        //{
        //    if (!File.Exists(path))
        //    {
        //        Console.WriteLine("没有找到文件");

        //    }
        //    IWorkbook workbook = null;

        //    using (FileStream fs = File.OpenRead(path))
        //    {
        //        workbook = new HSSFWorkbook(fs);




        //    }



        //    ISheet isheet = workbook.GetSheetAt(0);

        //    string sheetname = isheet.SheetName;
        //    int rowsNum = isheet.LastRowNum;
        //    IRow rw = null;
        //    int startRow = 0, endRow = 0;
        //    for (int i = 0; i < rowsNum; i++)
        //    {
        //        rw = isheet.GetRow(i);
        //        if (rw.Cells[0].ToString().Contains("序号"))
        //        {
        //            startRow = i + 1;
        //        }
        //        else if (rw.Cells[0].ToString().Contains("合计"))
        //        {
        //            endRow = i + 1;
        //            break;
        //        }
        //    }
        //    for (int i = startRow + 1; i < endRow-1; i++)
        //    {
        //        rw = isheet.GetRow(i);
        //        for (int j = 0; j < rw.Cells.Count; j++)
        //        {



        //            Console.WriteLine(rw.Cells[j]);






        //        }
        //    }
        //}

        /// <summary>
        /// 判断是否为数字
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static bool IsInteger(string s)




        {




            string pattern = @"^\d*$";




            return Regex.IsMatch(s, pattern);




        }

    }
}
