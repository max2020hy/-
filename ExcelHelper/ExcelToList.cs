using ExcelHelper;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHelper
{
    public static class operation
    {
        /// <summary>
        ///  根据模板创建新的文件
        /// </summary>
        /// <param name="lst"></param>
        public static List<Modle> ExcelToLis (string filePullPath,string start ,string end)
        {
            #region 加载xls文件

            //模板文件路径
            string path = filePullPath;
            FileStream FS;
            using (FS = new FileStream(path, FileMode.Open, FileAccess.ReadWrite))

            {


                IWorkbook workbook_model = new NPOI.HSSF.UserModel.HSSFWorkbook(FS);
                //  fs_modle.Close();

                ISheet sheet_model = workbook_model.GetSheetAt(0);


              //  IRow r = sheet_model.GetRow(row);//合计行
                int rnum = sheet_model.LastRowNum;
                int s = 0;int e=0;
                IRow r;
                for (int i = 0; i < rnum; i++)
                {
                    try
                    {
                        r = sheet_model.GetRow(i);
                        CellType type = r.GetCell(0).CellType;
                        //  Console.WriteLine(type.ToString());
                        string v = null;
                        switch (type)
                        {
                            case CellType.Unknown:
                                break;
                            case CellType.Numeric:
                                v = r.GetCell(0).NumericCellValue.ToString();
                                break;
                            case CellType.String:
                                v = r.GetCell(0).StringCellValue;
                                break;
                            case CellType.Formula:
                                break;
                            case CellType.Blank:
                                v = "";
                                break;
                            case CellType.Boolean:
                                break;
                            case CellType.Error:
                                break;
                            default:
                                break;
                        }
                        Console.WriteLine($"{v}*********{type}");
                        if (v == start)
                        {
                            s = i + 1;

                        }
                        else if (v == end)
                        {
                            e = i;
                            break;
                        }
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                 
                }
                int rowNum = e - s;
               

                List<Modle> lst= new List<Modle>();

                for (int i = s; i < e; i++)
                {
                    Modle m = new Modle();

                    m.Name = sheet_model.GetRow(i).GetCell(1).StringCellValue;
                    m.CarNum = sheet_model.GetRow(i).GetCell(2).StringCellValue;
                    m.Start = sheet_model.GetRow(i).GetCell(3).StringCellValue;
                    m.End = sheet_model.GetRow(i).GetCell(4).StringCellValue;

                    lst.Add(m);
                }

            return lst.Where(t=>t.Name!=string.Empty).ToList();
            }
            #endregion

            

        }
     
        public static void CreatFile(List<Modle> lst )
        {
            string value0 = "0002{0}~~停车费({1})~~~~~~0.05~~~~~~0~~False~~0000000000~~False~~30405020202~~否~~车辆停放服务~~~~~~35.0";
            string value1 = "102{0}~~{1}~~~~~~~~~~~~~~False";

         
            string SPBM = $"商品编码{DateTime.Now.Month}.txt";
            string KHBM= $"客户编码{DateTime.Now.Month}.txt";
            File.Copy("商品编码.txt", SPBM,true);
            File.Copy("客户编码.txt", KHBM,true);
         

            if (!File.Exists("客户编码.txt"))
            {
                return;
            }
           
                FileStream FS = new FileStream("客户编码.txt", FileMode.Open);
            
           
                Encoding type =  GetType(FS);
                Console.WriteLine(type);
            FS.Dispose();



            StreamWriter SW0 = new StreamWriter(SPBM, true, type);
            StreamWriter SW1 = new StreamWriter(KHBM, true, type);
            int i = 1;
            foreach (var m in lst)
            {
                string date0 = string.Format(value0,i.ToString().PadLeft(4,'0'), m.Start + "-" + m.End);
                SW0.WriteLine(date0);
                string date1 = string.Format(value1, i.ToString().PadLeft(4, '0'), m.Name + " " + m.CarNum);
                SW1.WriteLine(date1);
                i++;
            }
            SW0.Dispose();
            SW1.Dispose();
        }
        /// <summary> 
        /// 通过给定的文件流，判断文件的编码类型 
        /// </summary> 
        /// <param name=“fs“>文件流</param> 
        /// <returns>文件的编码类型</returns> 
        public static System.Text.Encoding GetType(FileStream fs)
        {
            byte[] Unicode = new byte[] { 0xFF, 0xFE, 0x41 };
            byte[] UnicodeBIG = new byte[] { 0xFE, 0xFF, 0x00 };
            byte[] UTF8 = new byte[] { 0xEF, 0xBB, 0xBF }; //带BOM 
            Encoding reVal = Encoding.Default;

            BinaryReader r = new BinaryReader(fs, System.Text.Encoding.Default);
            int i;
            int.TryParse(fs.Length.ToString(), out i);
            byte[] ss = r.ReadBytes(i);
            if ((ss[0] == 0xEF && ss[1] == 0xBB && ss[2] == 0xBF))
            {
                reVal = Encoding.UTF8;
            }
            else if (ss[0] == 0xFE && ss[1] == 0xFF && ss[2] == 0x00)
            {
                reVal = Encoding.BigEndianUnicode;
            }
            else if (ss[0] == 0xFF && ss[1] == 0xFE && ss[2] == 0x41)
            {
                reVal = Encoding.Unicode;
            }
            r.Close();
            return reVal;

        }

    }
}
