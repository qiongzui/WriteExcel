using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reflection;


namespace 计算EXCEL
{
    class Program
    {
        //写入EXCEL
        static public bool WriteXls(string filename)
        {
            //启动Excel应用程序
            Microsoft.Office.Interop.Excel.Application xls = new Microsoft.Office.Interop.Excel.Application();
            //_Workbook book = xls.Workbooks.Add(Missing.Value); //创建一张表，一张表可以包含多个sheet

            //如果表已经存在，可以用下面的命令打开
            _Workbook book = xls.Workbooks.Open(filename, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            _Worksheet sheet;//定义sheet变量

            int count = book.Sheets.Count;
            for( int n = 1; n < count + 1; n++)
            {
                try
                {
                    sheet = (_Worksheet)book.Worksheets.get_Item(n);//获得第i个sheet，准备写入
                }
                catch (Exception ex)//不存在就增加一个sheet
                {
                    return false;
                }

                //总行列数
                int rows = sheet.UsedRange.Rows.Count;
                int cols = sheet.UsedRange.Columns.Count;

                //公式所在第一行
                int row_first = 0;
                int col_fun = 0;
                int col_result = 0;

                for (int i = 1; i < rows; i++)
                {
                    //查找公式所在行列
                    if (row_first == 0)
                    {
                        for (int j = 1; j < cols; j++)
                        {
                            string value = sheet.Cells[i, j].Cells.Text.ToString();
                            if (value == "计算式")
                            {
                                row_first = i + 1;
                                col_fun = j;
                                col_result = j + 2;
                                break;
                            }

                        }
                    }

                    //开始处理
                    if (i >= row_first && row_first != 0)
                    {
                        string value = sheet.Cells[i, col_fun].Cells.Text.ToString();
                        //没有公式不处理
                        if ("计算式" == value || value == "")
                        {
                            continue;
                        }

                        MSScriptControl.ScriptControl sc = new MSScriptControl.ScriptControl();
                        sc.Language = "JavaScript";

                        string strN = value.Replace("（", "(");
                        value = strN.Replace("）", ")");
                        
                        string result = sc.Eval(value).ToString();
                        //设置单元格的值        
                        sheet.Cells[i, col_result] = result;
                    }
                }
            }
 
            
            //将表另存为
            //book.SaveAs(filename, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            //如果表已经存在，直接用下面的命令保存即可
            book.Save();

            book.Close(false, Missing.Value, Missing.Value);//关闭打开的表
            xls.Quit();//Excel程序退出
            //sheet,book,xls设置为null，防止内存泄露
            sheet = null;
            book = null;
            xls = null;
            GC.Collect();//系统回收资源
            return true;
        }

        static void Main(string[] args)
        {          
            Console.WriteLine("---------------开始处理----------------------");
            string strPath = System.Environment.CurrentDirectory;
            //string strPath = "F:\\MyProject";
            var files = Directory.GetFiles(strPath , "*.xls");
            if( files.Length != 0)
            {
                foreach(var file in files)
                {
                //    Array value = ReadXls(file);
                    Console.Write("正在处理：" + file + "...");
                    //写入结果
                    WriteXls(file);
                    Console.Write("完成!\n");
                }
            }
        }
    }
}
