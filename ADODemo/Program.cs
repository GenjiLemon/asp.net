using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ADODemo
{
    class Program
    {
        private static SqlConnection conn;
        private DataSet ds;
        private void InitConn()
        {
            String constr = "Server=112.74.104.17;Database=CMS;Uid=fay;Pwd=Root123;";
            conn = new SqlConnection(constr);
            string sql = "select * from user1";
            conn.Open();
            SqlCommand cmd = new SqlCommand(sql, conn);
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
             ds = new DataSet();
            sda.Fill(ds);
            conn.Close();
        }
        static void Main(string[] args)
        {
            // Program program = new Program();
            //program.InitConn();
            //program.ExcelTest();
            string wordpath=new word().test();
            string newpath = @"d:\1.pdf";
            new word2pdf().WordToPdf(wordpath, newpath);
        }
        public  void ExcelTest()
        {
            //导出：将数据库中的数据，存储到一个excel中

            //1、查询数据库数据  

            //2、  生成excel
            //2_1、生成workbook
            //2_2、生成sheet
            //2_3、遍历集合，生成行
            //2_4、根据对象生成单元格
            HSSFWorkbook workbook = new HSSFWorkbook();
            //创建工作表
            var sheet = workbook.CreateSheet("信息表");
            //创建标题行（重点）
            var row = sheet.CreateRow(0);
            //创建单元格
            var cellhead1 = row.CreateCell(0);
            cellhead1.SetCellValue("编号");
            var cellhead2 = row.CreateCell(1);
            cellhead2.SetCellValue("用户名");
  
            foreach(DataRow drow in ds.Tables[0].Rows)
            {
                var newrow = sheet.CreateRow(ds.Tables[0].Rows.IndexOf(drow));
                for(int i=0;i<ds.Tables[0].Columns.Count;i++)
                {    
                    var cellname = newrow.CreateCell(i);
                    cellname.SetCellValue(drow[ds.Tables[0].Columns[i]].ToString());
                }
            }

            FileStream file = new FileStream(@"D:\3.xls", FileMode.CreateNew, FileAccess.Write);
            workbook.Write(file);
            file.Dispose();
        }
        public void export2csv(DataSet ds)
        {
            using (StreamWriter streamWriter = new StreamWriter(@"d:\1.csv", false, Encoding.Default))
            {
                if (ds.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dataRow in ds.Tables[0].Rows)
                    {
                        StringBuilder str = new StringBuilder();
                        foreach (DataColumn dataColumn in ds.Tables[0].Columns)
                        {
                            str.Append(dataRow[dataColumn.Caption].ToString());
                            str.Append(",");
                        }
                        streamWriter.WriteLine(str);
                    }

                }
                else
                {
                    Console.WriteLine("没有数据输出!");
                }
                streamWriter.Flush();
                streamWriter.Close();
            }
        }
        public string CSVSaveasXLSX(string FilePath)
        {
            QuertExcel();
            string NewFilePath = "";

            Excel.Application excelApplication;
            Excel.Workbooks excelWorkBooks = null;
            Excel.Workbook excelWorkBook = null;
            Excel.Worksheet excelWorkSheet = null;

            try
            {
                excelApplication = new Excel.Application();
                excelWorkBooks = excelApplication.Workbooks;
                excelWorkBook = ((Excel.Workbook)excelWorkBooks.Open(FilePath, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value));
                excelWorkSheet = (Excel.Worksheet)excelWorkBook.Worksheets[1];
                excelApplication.Visible = false;
                excelApplication.DisplayAlerts = false;
                NewFilePath = FilePath.Replace(".csv", ".xlsx");
                excelWorkBook.SaveAs(NewFilePath, Excel.XlFileFormat.xlAddIn, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                excelWorkBook.Close();
                QuertExcel();
                GC.Collect(System.GC.GetGeneration(excelWorkSheet));
                GC.Collect(System.GC.GetGeneration(excelWorkBook));
                GC.Collect(System.GC.GetGeneration(excelApplication));
            }
            catch (Exception exc)
            {
                throw new Exception(exc.Message);
            }

            finally
            {
                GC.Collect();
            }
            return NewFilePath;
        }
        private void QuertExcel()
        {
            Process[] excels = Process.GetProcessesByName("EXCEL");
            foreach (var item in excels)
            {
                item.Kill();
            }
        }


    }


}


