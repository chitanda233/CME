using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using NPOI;
using NPOI.XSSF;
using NPOI.HSSF;
using NPOI.HSSF.Record;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming;
using NPOI.XSSF.UserModel;

namespace CMExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            string filename = "";
            if (args.Length!=1)
            {
                Console.WriteLine("使用说明：\nCMExcel.exe  表格名 \n" +
                                  "会将文件夹下文件名符合  表格名_*.xls* 的文件全部合并为 表格名.xlsx\n" +
                                  "在下方直接输入'表格名'(不包含后缀)也可执行一次\n" +
                                  "如果符合条件的文件不存在，则会创建 表格名.xlsx 空文件");
                while (true)
                {
                    filename=Console.ReadLine();
                    if (filename=="exit")
                    {
                        Environment.Exit(0);
                    }
                    if (filename!="")
                    {
                        break;
                    }
                }
            }
            else
                filename = args[0].ToString();
                
                

            DataTable tmpDataTable=new DataTable();
            DataTable outDataTable=new DataTable();
            int time = 1;

            CMETools cm=new CMETools();

            var foInfo = new DirectoryInfo(System.Environment.CurrentDirectory);
            foreach (var file in foInfo.GetFiles(filename+"_*.xls*"))
            {
                if (file==null)
                {
                    return;
                }
                tmpDataTable = cm.ExcelToDataTable(file.ToString(), 0, true);
                if (time==1)
                {
                    outDataTable = tmpDataTable.Clone();
                    outDataTable.Rows.Clear();
                    object[] objc = new object[outDataTable.Columns.Count];
                    for (int i = 0; i < tmpDataTable.Rows.Count; i++)
                    {
                        tmpDataTable.Rows[i].ItemArray.CopyTo(objc,0);
                        outDataTable.Rows.Add(objc);
                    }
                    time += 1;
                }
                else
                {
                    object[] objc = new object[outDataTable.Columns.Count];
                    for (int i = 0; i < tmpDataTable.Rows.Count; i++)
                    {
                        tmpDataTable.Rows[i].ItemArray.CopyTo(objc, 0);
                        outDataTable.Rows.Add(objc);
                    }
                    time += 1;
                }
            }

            List<DataTable> dl=new List<DataTable>();
            dl.Add(outDataTable);
            cm.DataTableToExcel(filename+".xlsx", dl, true);
           
        }

        static void WriteToFile(IWorkbook iw)
        {
            //Write the stream data of workbook to the root directory
            FileStream file = new FileStream(@"test.xls", FileMode.OpenOrCreate);
            iw.Write(file);
            file.Close();
        }

        
    }

    public class CMETools
    {
 
        /// <summary>
        /// 从一个excel中读取为dataTable
        /// </summary>
        /// <param name="fileName">excel文件名</param>
        /// <param name="sheetPosition">读第几个页签,从0开始</param>
        /// <param name="isFirstRowColumn">是不是从第一行读取，用来初始化dt的结构</param>
        /// <returns></returns>
        public DataTable ExcelToDataTable(string fileName, int sheetPosition, bool isFirstRowColumn)
        {
            ISheet sheet = null;
            DataTable data = new DataTable();
            int startRow = 0;
            IWorkbook workbook = null;

            try
            {
                //读取文件并判断格式
                var fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                if (fileName.IndexOf(".xlsx", StringComparison.Ordinal) > 0) // 2007版本
                    workbook = new XSSFWorkbook(fs);
                else if (fileName.IndexOf(".xls", StringComparison.Ordinal) > 0) // 2003版本
                    workbook = new HSSFWorkbook(fs);
                else
                {
                    Console.WriteLine("no excel file");
                    return null;
                }

                //读指定sheet
                sheet = workbook.GetSheetAt(sheetPosition);
                if (sheet != null)
                {
                    IRow firstRow = sheet.GetRow(0);
                    startRow = sheet.FirstRowNum;

                    //获取当前表格最大列数(表头)
                    int cellCount = 0;
                    for (int i = 0; i < firstRow.LastCellNum; i++)
                    {
                        if (firstRow.GetCell(i).CellType != CellType.Blank)
                        {
                            cellCount++;
                        }
                    }

                    if (isFirstRowColumn)
                    {
                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                        {
                            ICell cell = firstRow.GetCell(i);
                            if (cell != null)
                            {
                                string cellValue = cell.StringCellValue;
                                if (cellValue != null)
                                {
                                    DataColumn column = new DataColumn(cellValue);
                                    data.Columns.Add(column);
                                }
                            }
                        }
                        startRow = sheet.FirstRowNum + 1;
                    }
                    else
                    {
                        startRow = sheet.FirstRowNum;
                    }


                    //最后一列的标号
                    int rowCount = sheet.LastRowNum;

                    for (int i = startRow; i <= rowCount; i++)
                    {

                        IRow row = sheet.GetRow(i);
                        if (row == null) continue; //没有数据的行默认是null　　

                        DataRow dataRow = data.NewRow();

                        for (int j = 0; j < cellCount; j++)
                        {
                            if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null
                            {
                                row.GetCell(j).SetCellType(CellType.String);
                                dataRow[j] = row.GetCell(j).StringCellValue;
                            }
                        }
                        data.Rows.Add(dataRow);
                    }
                }
                fs.Close();
                return data;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return null;
            }
        }

        public int DataTableToExcel(string outPutName, List<DataTable> dataList, bool isColumnWritten)
        {
            int i = 0;
            int j = 0;
            int count = 0;
            ISheet sheet = null;
            string sheetName = null;
            string all = null;
            IWorkbook workbook = null;

            var fs = new FileStream(outPutName, FileMode.Create, FileAccess.Write);
            workbook = new XSSFWorkbook();

            for (int m = 0; m < dataList.Count; m++)
            {

                if (m >= 3)
                {
                    return -1;
                }

                switch (m)
                {
                    case 0: sheetName = "sheet1"; break;
                    case 1: sheetName = "sheet2"; break;
                    case 2: sheetName = "sheet3"; break;
                }
                DataTable data = dataList[m];
                sheet = workbook.CreateSheet(sheetName);


                if (isColumnWritten == true) //写入DataTable的列名
                {
                    IRow row = sheet.CreateRow(0);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Columns[j].ColumnName);
                    }
                    count = 1;
                }
                else
                {
                    count = 0;
                }

                for (i = 0; i < data.Rows.Count; ++i)
                {
                    IRow row = sheet.CreateRow(count);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Rows[i][j].ToString());
                    }
                    ++count;
                }
            }
            try
            {


                workbook.Write(fs); //写入到excel
                fs.Close();
                fs = null;
                return count;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                Console.ReadKey();
                return -1;
            }
        }
    }
}
