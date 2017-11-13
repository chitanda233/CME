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
            DataTable x=ExcelToDataTable("test_name1.xlsx", 0, true);
            List<DataTable> dl=new List<DataTable>();
            dl.Add(x);
            DataTableToExcel("hahah.xlsx", dl, true);
            Console.ReadKey();
        }

        static void WriteToFile(IWorkbook iw)
        {
            //Write the stream data of workbook to the root directory
            FileStream file = new FileStream(@"test.xls", FileMode.OpenOrCreate);
            iw.Write(file);
            file.Close();
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="sheetPosition"></param>
        /// <param name="isFirstRowColumn"></param>
        /// <returns></returns>
       static public DataTable ExcelToDataTable(string fileName, int sheetPosition, bool isFirstRowColumn)
        {
            ISheet sheet = null;
            DataTable data = new DataTable();
            
            int startRow = 0;
            IWorkbook workbook=null;
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

        static  public int DataTableToExcel(string outPutName, List<DataTable> dataList, bool isColumnWritten)
        {
            int i = 0;
            int j = 0;
            int count = 0;
            ISheet sheet = null;
            string sheetName = null;
            string all = null;
            IWorkbook workbook=null;

            var fs = new FileStream(outPutName, FileMode.Append, FileAccess.Write);
            workbook = new XSSFWorkbook();
          
            for (int m = 0; m < dataList.Count; m++)
            {
                //for (int k = 0; k < dataList[m].Rows.Count; k++) {
                //    for (int p = 1; p < dataList[m].Columns.Count; p++) {
                //        if (dataList[m].Rows[k][p].ToString() == "")
                //        {
                //            Console.WriteLine();
                //        }
                //        all += dataList[m].Rows[k][p].ToString();
                //        if (all=="") {
                //            dataList[m].Rows.Remove(dataList[m].Rows[k]);
                //        }
                //    }
                //}
                switch (m)
                {
                    case 0: sheetName = "sheet1"; break;
                    case 1: sheetName = "表2机构网络情况调查表"; break;
                    case 2: sheetName = "表3机构硬盘录像机设备调查表"; break;
                    case 3: sheetName = "表4机构报警主机设备调查表"; break;
                    case 4: sheetName = "表5摄像机与硬盘录像机通道对应关系统计表"; break;
                    case 5: sheetName = "表6报警器、硬盘录像机、联动摄像机对应关系调查表"; break;
                    case 6: sheetName = "表7报警器、报警主机、联动摄像机对应关系调查表"; break;
                    case 7: sheetName = "表8机构门禁设备调查表"; break;
                    case 8: sheetName = "表9机构对讲设备调查表"; break;
                }
                DataTable data = dataList[m];
                if (workbook != null)
                {
                    sheet = workbook.CreateSheet(sheetName);

                }
                else
                {
                    return -1;
                }

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
                return count;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return -1;
            }
        }


    }
}
