using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;

namespace FileHelper
{
    /// <summary>
    /// Excel帮助类
    /// </summary>
    public class ExcelHelper
    {

        /// <summary>
        /// 将excel导入到datatable
        /// </summary>
        /// <param name="filePath">excel路径</param>
        /// <param name="isColumnName">第一行是否是列名</param>
        /// <returns>返回datatable</returns>
        public static DataTable ExcelToDataTable(string filePath, bool isColumnName)
        {
            DataTable dataTable = null;
            FileStream fs = null;
            DataColumn column = null;
            DataRow dataRow = null;
            IWorkbook workbook = null;
            ISheet sheet = null;
            IRow row = null;
            ICell cell = null;
            int startRow = 0;//内容起始行（不包含列名所在行）
            int SheetCount = 1;//Sheet总数
            try
            {
                using (fs = File.OpenRead(filePath))
                {
                    // 2007版本
                    if (filePath.IndexOf(".xlsx") > 0)
                        workbook = new XSSFWorkbook(fs);
                    // 2003版本
                    else if (filePath.IndexOf(".xls") > 0)
                        workbook = new HSSFWorkbook(fs);

                    if (workbook != null)
                    {
                        SheetCount = workbook.NumberOfSheets;//sheet数量
                        sheet = workbook.GetSheetAt(0);//读取第一个sheet，当然也可以循环读取每个sheet

                        dataTable = new DataTable();
                        if (sheet != null)
                        {
                            int rowCount = sheet.LastRowNum;//总行数
                            if (rowCount > 0)
                            {
                                IRow firstRow = sheet.GetRow(0);//第一行
                                int cellCount = firstRow.LastCellNum;//列数

                                //构建datatable的列
                                if (isColumnName)
                                {
                                    startRow = 1;//如果第一行是列名，则从第二行开始读取
                                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                                    {
                                        cell = firstRow.GetCell(i);
                                        if (cell != null)
                                        {
                                            if (cell.StringCellValue != null)
                                            {
                                                column = new DataColumn(cell.StringCellValue);
                                                dataTable.Columns.Add(column);
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                                    {
                                        column = new DataColumn("column" + (i + 1));
                                        dataTable.Columns.Add(column);
                                    }
                                }

                                //填充行
                                for (int i = startRow; i <= rowCount; ++i)
                                {
                                    row = sheet.GetRow(i);
                                    if (row == null)
                                        continue;

                                    dataRow = dataTable.NewRow();
                                    for (int j = row.FirstCellNum; j < cellCount; ++j)
                                    {
                                        cell = row.GetCell(j);
                                        if (cell == null)
                                        {
                                            dataRow[j] = "";
                                        }
                                        else
                                        {
                                            //CellType(Unknown = -1,Numeric = 0,String = 1,Formula = 2,Blank = 3,Boolean = 4,Error = 5,)
                                            switch (cell.CellType)
                                            {
                                                case CellType.Blank:
                                                    dataRow[j] = "";
                                                    break;
                                                case CellType.Numeric:
                                                    short format = cell.CellStyle.DataFormat;
                                                    //对时间格式（2015.12.5、2015/12/5、2015-12-5等）的处理
                                                    if (format == 14 || format == 31 || format == 57 || format == 58)
                                                        dataRow[j] = cell.DateCellValue;
                                                    else
                                                        dataRow[j] = cell.NumericCellValue;
                                                    break;
                                                case CellType.String:
                                                    dataRow[j] = cell.StringCellValue;
                                                    break;
                                            }
                                        }
                                    }
                                    dataTable.Rows.Add(dataRow);
                                }
                            }
                        }
                    }
                }
                return dataTable;
            }
            catch (Exception ex)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                return null;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath"></param>
        public static void SplitWorkBookSheets(string filePath)
        {
            FileStream fs = null;
            IWorkbook workbook = null;
            ISheet sheet = null;
            IRow row = null;
            ICell cell = null;
            int SheetCount = 0;
            int currentIndex = 0;
            List<int> SaveIndex = new List<int>();
            do
            {
                fs = File.OpenRead(filePath);
                // 2007版本
                if (filePath.IndexOf(".xlsx") > 0)
                    workbook = new XSSFWorkbook(fs);
                // 2003版本
                else if (filePath.IndexOf(".xls") > 0)
                    workbook = new HSSFWorkbook(fs);

                fs.Close();
                fs.Dispose();

                if (workbook != null)
                {
                    SheetCount = workbook.NumberOfSheets - currentIndex;//sheet数量
                    //删除后面的
                    for (int i = 0; i < SheetCount - 1; i++)
                    {
                        workbook.RemoveSheetAt(currentIndex + 1);
                    }
                    //删除前面的
                    foreach (var item in SaveIndex)
                    {
                        workbook.RemoveSheetAt(0);
                    }
                    //另存为文件
                    MemoryStream ms = new MemoryStream();
                    workbook.Write(ms);
                    FileStream fs2 = new FileStream("E:\\" + workbook.GetSheetName(0) + ".xlsx", FileMode.Create);
                    fs2.Write(ms.ToArray(), 0, ms.ToArray().Length);
                    ms.Close();
                    fs2.Close();
                }
                workbook.Close();
                SaveIndex.Add(currentIndex);
                currentIndex++;

            } while (currentIndex < SheetCount);


        }

        /// <summary>
        /// 导出excel至文件流
        /// </summary>
        /// <returns></returns>
        public static Stream ExportExcelByList<T>(List<T> list, string fileName)
        {
            MemoryStream ms = new MemoryStream();
            IWorkbook workbook = new HSSFWorkbook();
            //创建sheet
            ISheet sheet = workbook.CreateSheet(fileName);
            //数据
            Type type = typeof(T);
            PropertyInfo[] pis = type.GetProperties();
            int pisLen = pis.Length;
            //获取表头
            List<ExcelHeader> piColumn = GetHeaderColumns(type, pis);
            //表头样式
            ICellStyle style = SetStyle(workbook);
            //填充
            FillHeader(piColumn, sheet, style);

            ICellStyle bodyStyle = workbook.CreateCellStyle();
            bodyStyle.BorderBottom = BorderStyle.Thin;
            bodyStyle.BorderLeft = BorderStyle.Thin;
            bodyStyle.BorderRight = BorderStyle.Thin;
            bodyStyle.BorderTop = BorderStyle.Thin;
            //填充
            FillBody<T>(list, piColumn, pis, sheet, bodyStyle);

            //写入
            workbook.Write(ms);
            ms.Flush();
            ms.Position = 0;
            workbook = null;
            return ms;
        }

        /// <summary>
        /// 保存至本地
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="path"></param>
        public static void Save(Stream stream, string path)
        {
            byte[] srcBuf = new Byte[stream.Length];
            stream.Read(srcBuf, 0, srcBuf.Length);
            stream.Seek(0, SeekOrigin.Begin);
            stream.Flush();
            stream.Position = 0;
            stream.Close();
            //判断路径是否正确
            if (!string.IsNullOrEmpty(path))
            {
                try
                {
                    using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
                    {
                        fs.Write(srcBuf, 0, srcBuf.Length);
                        fs.Close();
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        /// <summary>
        /// 获取表头
        /// </summary>
        /// <param name="propertyArray"></param>
        /// <returns></returns>
        private static List<ExcelHeader> GetHeaderColumns(Type type, PropertyInfo[] propertyArray)
        {
            List<ExcelHeader> HeaderColumns = new List<ExcelHeader>();
            if (propertyArray == null)
            {
                return HeaderColumns;
            }

            for (int i = 0; i < propertyArray.Length; i++)
            {
                PropertyInfo propertyItem = propertyArray[i];
                string displayName = GetDisplayName(type, propertyItem.Name);
                if (!string.IsNullOrEmpty(displayName))
                {
                    HeaderColumns.Add(new ExcelHeader
                    {
                        Index = i,
                        Name = GetDisplayName(type, propertyArray[i].Name),
                        Display = true
                    });
                }
                else
                {
                    HeaderColumns.Add(new ExcelHeader
                    {
                        Index = i,
                        Name = GetDisplayName(type, propertyArray[i].Name),
                        Display = false
                    });
                }
            }

            return HeaderColumns;
        }

        /// <summary>
        /// 设置样式
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        private static ICellStyle SetStyle(IWorkbook workbook)
        {
            //样式
            ICellStyle style = workbook.CreateCellStyle();
            //设置单元格的样式：水平对齐居中
            style.VerticalAlignment = VerticalAlignment.Center;
            style.Alignment = HorizontalAlignment.Center;
            //style.FillForegroundColor = HSSFColor.Automatic.Index;
            //style.FillPattern = FillPattern.SolidForeground;
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;

            //新建一个字体样式对象
            IFont font = workbook.CreateFont();
            font.FontName = "宋体";
            font.FontHeightInPoints = 11;
            //设置字体加粗样式
            font.Boldweight = (short)FontBoldWeight.Bold;
            //使用SetFont方法将字体样式添加到单元格样式中 
            style.SetFont(font);

            return style;
        }

        /// <summary>
        /// 填充表头
        /// </summary>
        /// <param name="columns"></param>
        /// <param name="sheet"></param>
        /// <param name="style"></param>
        /// <returns></returns>
        private static int FillHeader(List<ExcelHeader> columns, ISheet sheet, ICellStyle style)
        {
            if (columns == null)
            {
                return 0;
            }
            int colIndex = 0;
            //创建表头行
            IRow headerRow = sheet.CreateRow(0);
            //行高
            headerRow.HeightInPoints = 35;
            //填充
            foreach (var item in columns)
            {
                if (item.Display)
                {
                    sheet.SetColumnWidth(colIndex, 22 * 256 + 200);
                    var Cell = headerRow.CreateCell(colIndex);
                    Cell.SetCellValue(item.Name);
                    //将新的样式赋给单元格
                    Cell.CellStyle = style;
                    colIndex++;
                }
            }
            return colIndex;
        }

        /// <summary>
        /// 填充数据行
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list"></param>
        /// <param name="sheet"></param>
        private static void FillBody<T>(List<T> list, List<ExcelHeader> columns, PropertyInfo[] propertyArray, ISheet sheet, ICellStyle style)
        {
            if (list == null || list.Count < 1)
            {
                return;
            }

            int rowIndex = 1;
            //行
            foreach (T rowItem in list)
            {
                IRow dataRow = sheet.CreateRow(rowIndex);
                int colIndex = 0;
                //列
                foreach (var colItem in columns)
                {
                    if (colItem.Display == false)
                    {
                        continue;
                    }
                    try
                    {
                        PropertyInfo pi = propertyArray[colItem.Index];
                        object objVal = pi.GetValue(rowItem);
                        string value = "";
                        if (objVal != null)
                        {
                            Type type = objVal.GetType();
                            if (type.IsValueType || type == typeof(string))
                            {
                                value = objVal.ToString();
                            }
                        }

                        var Cell = dataRow.CreateCell(colIndex);
                        Cell.SetCellValue(value);
                        Cell.CellStyle = style;
                    }
                    catch (Exception ex)
                    {
                        var Cell = dataRow.CreateCell(colIndex);
                        Cell.SetCellValue("Error:数据异常");
                        Cell.CellStyle = style;
                    }
                    colIndex++;
                }
                rowIndex++;
            }
        }

        /// <summary>
        /// 值转换
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <param name="pi"></param>
        /// <returns></returns>
        public static string ConvertType<T>(T data, PropertyInfo pi)
        {
            Type type = pi.GetValue(data).GetType();
            //值类型
            if (type.IsValueType)
            {
                return pi.GetValue(data).ToString();
            }
            //引用类型string
            else if (type == typeof(string))
            {
                return pi.GetValue(data).ToString();
            }
            //引用类型
            else
            {
                return "";
            }
        }
        /// <summary>
        /// 获取模型中DisplayName的值
        /// </summary>
        /// <param name="modelType">类型</param>
        /// <param name="propertyDisplayName">成员名称</param>
        /// <returns></returns>
        public static string GetDisplayName(Type modelType, string propertyName)
        {
            return (System.ComponentModel.TypeDescriptor.GetProperties(modelType)[propertyName].Attributes[typeof(System.ComponentModel.DisplayNameAttribute)] as System.ComponentModel.DisplayNameAttribute).DisplayName;
        }
    }
    /// <summary>
    /// 
    /// </summary>
    public class ExcelHeader
    {
        /// <summary>
        /// 列索引
        /// </summary>
        public int Index { get; set; }
        /// <summary>
        /// 列明
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 是否显示
        /// </summary>
        public bool Display { get; set; }
    }
}
