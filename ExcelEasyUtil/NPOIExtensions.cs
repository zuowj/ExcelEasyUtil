using Newtonsoft.Json.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq.Expressions;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text;

namespace ExcelEasyUtil
{

    /// <summary>
    /// NPOI扩展类
    /// author:zuowenjun 
    /// 2019-5-21
    /// </summary>
    public static class NPOIExtensions
    {
        /// <summary>
        /// 将一个实体数据对象填充到一个EXCEL工作表中(可连续填充多个sheet，如：FillSheet(...).FillSheet(..) )
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="book"></param>
        /// <param name="sheetName"></param>
        /// <param name="headerColNames"></param>
        /// <param name="excelData"></param>
        /// <param name="getCellValueFunc"></param>
        /// <returns></returns>
        public static IWorkbook FillSheet<T>(this IWorkbook book, string sheetName, IList<T> excelData,
            IList<string> headerColNames, Func<T, List<object>> getCellValuesFunc, IDictionary<string, string> colDataFormats = null) where T : class
        {
            var sheet = book.CreateSheet(sheetName);

            IRow rowHeader = sheet.CreateRow(0);
            var headerCellStyle = GetCellStyle(book, true);
            Dictionary<int, ICellStyle> colStyles = new Dictionary<int, ICellStyle>();
            List<Type> colTypes = new List<Type>();
            Type strType = typeof(string);
            for (int i = 0; i < headerColNames.Count; i++)
            {
                ICell headerCell = rowHeader.CreateCell(i);
                headerCell.CellStyle = headerCellStyle;

                string colName = headerColNames[i];

                if (colName.Contains(":"))
                {
                    var colInfos = colName.Split(':');
                    colName = colInfos[0];
                    colTypes.Add(GetColType(colInfos[1]));
                }
                else
                {
                    colTypes.Add(strType);
                }

                headerCell.SetCellValue(colName);
                if (colDataFormats != null && colDataFormats.ContainsKey(colName))
                {
                    colStyles[i] = GetCellStyleWithDataFormat(book, colDataFormats[colName]);
                }
                else
                {
                    colStyles[i] = GetCellStyle(book);
                }
            }

            //生成excel内容
            for (int i = 0; i < excelData.Count; i++)
            {
                IRow irow = sheet.CreateRow(i + 1);
                var row = excelData[i];
                var cellValues = getCellValuesFunc(row);
                for (int j = 0; j < headerColNames.Count; j++)
                {
                    ICell cell = irow.CreateCell(j);
                    string cellValue = string.Empty;
                    if (cellValues.Count - 1 >= j && cellValues[j] != null)
                    {
                        cellValue = cellValues[j].ToString();
                    }

                    SetCellValue(cell, cellValue, colTypes[j], colStyles);
                }
            }

            return book;
        }

        /// <summary>
        ///  将一个实体数据对象填充到一个EXCEL工作表中(可连续填充多个sheet，如：FillSheet(...).FillSheet(..) )
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="book"></param>
        /// <param name="sheetName"></param>
        /// <param name="colMaps"></param>
        /// <param name="excelData"></param>
        /// <returns></returns>
        public static IWorkbook FillSheet<T>(this IWorkbook book, string sheetName, IList<T> excelData,
                IDictionary<string, Expression<Func<T, dynamic>>> colMaps, IDictionary<string, string> colDataFormats = null
            ) where T : class
        {

            var sheet = book.CreateSheet(sheetName);

            var headerColNames = new List<string>();
            var propInfos = new List<PropertyInfo>();

            foreach (var item in colMaps)
            {
                headerColNames.Add(item.Key);
                var propInfo = GetPropertyInfo(item.Value);
                propInfos.Add(propInfo);
            }

            var headerCellStyle = GetCellStyle(book, true);
            Dictionary<int, ICellStyle> colStyles = new Dictionary<int, ICellStyle>();
            IRow rowHeader = sheet.CreateRow(0);
            for (int i = 0; i < headerColNames.Count; i++)
            {
                ICell headerCell = rowHeader.CreateCell(i);
                headerCell.CellStyle = headerCellStyle;
                headerCell.SetCellValue(headerColNames[i]);

                if (colDataFormats != null && colDataFormats.ContainsKey(headerColNames[i]))
                {
                    colStyles[i] = GetCellStyleWithDataFormat(book, colDataFormats[headerColNames[i]]);
                }
                else
                {
                    colStyles[i] = GetCellStyle(book);
                }
            }

            //生成excel内容
            for (int i = 0; i < excelData.Count; i++)
            {
                IRow irow = sheet.CreateRow(i + 1);
                var row = excelData[i];
                for (int j = 0; j < headerColNames.Count; j++)
                {
                    var prop = propInfos[j];

                    ICell cell = irow.CreateCell(j);
                    string cellValue = Convert.ToString(propInfos[j].GetValue(row, null));
                    SetCellValue(cell, cellValue, prop.PropertyType, colStyles);
                }
            }

            return book;
        }

        /// <summary>
        /// 将一个数据表（DataTable）对象填充到一个EXCEL工作表中(可连续填充多个sheet，如：FillSheet(...).FillSheet(..) )
        /// </summary>
        /// <param name="book"></param>
        /// <param name="sheetName"></param>
        /// <param name="excelData"></param>
        /// <param name="colMaps"></param>
        /// <param name="colDataFormats"></param>
        /// <returns></returns>
        public static IWorkbook FillSheet(this IWorkbook book, string sheetName, DataTable excelData, IDictionary<string, string> colMaps,
            IDictionary<string, string> colDataFormats = null)
        {

            if (excelData.Rows.Count <= 0) return book;

            var sheet = book.CreateSheet(sheetName);


            var headerCellStyle = GetCellStyle(book, true);
            Dictionary<int, ICellStyle> colStyles = new Dictionary<int, ICellStyle>();
            IRow rowHeader = sheet.CreateRow(0);

            int nIndex = 0;
            var headerColNames = new List<string>();

            foreach (var item in colMaps)
            {
                ICell headerCell = rowHeader.CreateCell(nIndex);
                headerCell.SetCellValue(item.Value);
                headerCell.CellStyle = headerCellStyle;

                if (colDataFormats != null && colDataFormats.ContainsKey(item.Value))
                {
                    colStyles[nIndex] = GetCellStyleWithDataFormat(book, colDataFormats[item.Value]);
                }
                else
                {
                    colStyles[nIndex] = GetCellStyle(book);
                }
                headerColNames.Add(item.Key);
                nIndex++;
            }

            //生成excel内容
            for (int i = 0; i < excelData.Rows.Count; i++)
            {
                IRow irow = sheet.CreateRow(i + 1);
                var row = excelData.Rows[i];
                for (int j = 0; j < headerColNames.Count; j++)
                {
                    ICell cell = irow.CreateCell(j);
                    string colName = headerColNames[j];
                    string cellValue = row[colName].ToNotNullString();
                    SetCellValue(cell, cellValue, excelData.Columns[colName].DataType, colStyles);
                }
            }

            return book;
        }

        private static PropertyInfo GetPropertyInfo<T, TR>(Expression<Func<T, TR>> select)
        {
            var body = select.Body;
            if (body.NodeType == ExpressionType.Convert)
            {
                var o = (body as UnaryExpression).Operand;
                return (o as MemberExpression).Member as PropertyInfo;
            }
            else if (body.NodeType == ExpressionType.MemberAccess)
            {
                return (body as MemberExpression).Member as PropertyInfo;
            }
            return null;
        }

        private static Type GetColType(string colTypeSimpleName)
        {
            colTypeSimpleName = colTypeSimpleName.ToUpper();
            switch (colTypeSimpleName)
            {
                case "DT":
                    {
                        return typeof(DateTime);
                    }
                case "BL":
                    {
                        return typeof(Boolean);
                    }
                case "NUM":
                    {
                        return typeof(Int64);
                    }
                case "DEC":
                    {
                        return typeof(Decimal);
                    }
                default:
                    {
                        return typeof(String);
                    }
            }
        }


        private static void SetCellValue(ICell cell, string value, Type colType, IDictionary<int, ICellStyle> colStyles)
        {
            if (colType.IsNullableType())
            {
                colType = colType.GetGenericArguments()[0];
            }

            string dataFormatStr = null;
            switch (colType.ToString())
            {
                case "System.String": //字符串类型
                    cell.SetCellType(CellType.String);
                    cell.SetCellValue(value);
                    break;
                case "System.DateTime": //日期类型
                    DateTime dateV = new DateTime();
                    if (DateTime.TryParse(value, out dateV))
                    {
                        cell.SetCellValue(dateV);
                    }
                    else
                    {
                        cell.SetCellValue(value);
                    }
                    dataFormatStr = "yyyy/mm/dd hh:mm:ss";
                    break;
                case "System.Boolean": //布尔型
                    bool boolV = false;
                    if (bool.TryParse(value, out boolV))
                    {
                        cell.SetCellType(CellType.Boolean);
                        cell.SetCellValue(boolV);
                    }
                    else
                    {
                        cell.SetCellValue(value);
                    }
                    break;
                case "System.Int16": //整型
                case "System.Int32":
                case "System.Int64":
                case "System.Byte":
                    int intV = 0;
                    if (int.TryParse(value, out intV))
                    {
                        cell.SetCellType(CellType.Numeric);
                        cell.SetCellValue(intV);
                    }
                    else
                    {
                        cell.SetCellValue(value);
                    }
                    dataFormatStr = "0";
                    break;
                case "System.Decimal": //浮点型
                case "System.Double":
                    double doubV = 0;
                    if (double.TryParse(value, out doubV))
                    {
                        cell.SetCellType(CellType.Numeric);
                        cell.SetCellValue(doubV);
                    }
                    else
                    {
                        cell.SetCellValue(value);
                    }
                    dataFormatStr = "0.00";
                    break;
                case "System.DBNull": //空值处理
                    cell.SetCellType(CellType.Blank);
                    cell.SetCellValue(string.Empty);
                    break;
                default:
                    cell.SetCellType(CellType.Unknown);
                    cell.SetCellValue(value);
                    break;
            }

            if (!string.IsNullOrEmpty(dataFormatStr) && colStyles[cell.ColumnIndex].DataFormat <= 0) //没有设置，则采用默认类型格式
            {
                colStyles[cell.ColumnIndex] = GetCellStyleWithDataFormat(cell.Sheet.Workbook, dataFormatStr);
            }
            cell.CellStyle = colStyles[cell.ColumnIndex];

            ReSizeColumnWidth(cell.Sheet, cell);
        }


        private static ICellStyle GetCellStyleWithDataFormat(IWorkbook workbook, string format)
        {
            ICellStyle style = GetCellStyle(workbook);

            var dataFormat = workbook.CreateDataFormat();
            short formatId = -1;
            if (dataFormat is HSSFDataFormat)
            {
                formatId = HSSFDataFormat.GetBuiltinFormat(format);
            }
            if (formatId != -1)
            {
                style.DataFormat = formatId;
            }
            else
            {
                style.DataFormat = dataFormat.GetFormat(format);
            }
            return style;
        }


        private static ICellStyle GetCellStyle(IWorkbook workbook, bool isHeaderRow = false)
        {
            ICellStyle style = workbook.CreateCellStyle();

            if (isHeaderRow)
            {
                style.FillPattern = FillPattern.SolidForeground;
                style.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Grey25Percent.Index;
                IFont f = workbook.CreateFont();
                f.FontHeightInPoints = 11D;
                f.Boldweight = (short)FontBoldWeight.Bold;
                style.SetFont(f);
            }

            style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            return style;
        }

        private static void ReSizeColumnWidth(ISheet sheet, ICell cell)
        {
            int cellLength = (Encoding.Default.GetBytes(cell.ToString()).Length + 2) * 256;
            const int maxLength = 60 * 256; //255 * 256;
            if (cellLength > maxLength) //当单元格内容超过30个中文字符（英语60个字符）宽度，则强制换行
            {
                cellLength = maxLength;
                cell.CellStyle.WrapText = true;
            }
            int colWidth = sheet.GetColumnWidth(cell.ColumnIndex);
            if (colWidth < cellLength)
            {
                sheet.SetColumnWidth(cell.ColumnIndex, cellLength);
            }
        }

        private static ISheet GetSheet(IWorkbook workbook, string sheetIndexOrName)
        {
            int sheetIndex = 0;
            ISheet sheet = null;

            if (int.TryParse(sheetIndexOrName, out sheetIndex))
            {
                sheet = workbook.GetSheetAt(sheetIndex);
            }
            else
            {
                sheet = workbook.GetSheet(sheetIndexOrName);
            }
            return sheet;
        }

        /// <summary>
        /// 从工作表中解析生成DataTable
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheetIndexOrName"></param>
        /// <param name="headerRowIndex"></param>
        /// <param name="startColIndex"></param>
        /// <param name="colCount"></param>
        /// <returns></returns>
        public static DataTable ResolveDataTable(this IWorkbook workbook, string sheetIndexOrName, int headerRowIndex, short startColIndex = 0, short colCount = 0)
        {
            DataTable table = new DataTable();

            ISheet sheet = GetSheet(workbook, sheetIndexOrName);

            IRow headerRow = sheet.GetRow(headerRowIndex);
            int cellFirstNum = (startColIndex > headerRow.FirstCellNum ? startColIndex : headerRow.FirstCellNum);
            int cellCount = (colCount > 0 && colCount < headerRow.LastCellNum ? colCount : headerRow.LastCellNum);

            for (int i = cellFirstNum; i < cellCount; i++)
            {
                if (headerRow.GetCell(i) == null || headerRow.GetCell(i).StringCellValue.Trim() == "")
                {
                    // 如果遇到第一个空列，则不再继续向后读取
                    cellCount = i;
                    break;
                }
                DataColumn column = new DataColumn(headerRow.GetCell(i).StringCellValue);
                table.Columns.Add(column);
            }

            for (int i = (headerRowIndex + 1); i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    List<string> cellValues = new List<string>();
                    for (int j = cellFirstNum; j < cellCount; j++)
                    {
                        if (row.GetCell(j) != null)
                        {
                            cellValues.Add(row.GetCell(j).ToNotNullString());
                        }
                        else
                        {
                            cellValues.Add(string.Empty);
                        }
                    }

                    table.Rows.Add(cellValues.ToArray());
                }
            }

            return table;
        }

        /// <summary>
        /// 从工作表中解析生成指定的结果对象列表
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="workbook"></param>
        /// <param name="sheetIndexOrName"></param>
        /// <param name="headerRowIndex"></param>
        /// <param name="buildResultItemFunc"></param>
        /// <param name="startColIndex"></param>
        /// <param name="colCount"></param>
        /// <returns></returns>
        public static List<T> ResolveAs<T>(this IWorkbook workbook, string sheetIndexOrName, int headerRowIndex, Func<List<string>, T> buildResultItemFunc,
            short startColIndex = 0, short colCount = 0)
        {
            ISheet sheet = GetSheet(workbook, sheetIndexOrName);

            IRow headerRow = sheet.GetRow(headerRowIndex);
            int cellFirstNum = (startColIndex > headerRow.FirstCellNum ? startColIndex : headerRow.FirstCellNum);
            int cellCount = (colCount > 0 && colCount < headerRow.LastCellNum ? colCount : headerRow.LastCellNum);

            List<T> resultList = new List<T>();
            for (int i = (headerRowIndex + 1); i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    List<string> cellValues = new List<string>();
                    for (int j = cellFirstNum; j < cellCount; j++)
                    {
                        if (row.GetCell(j) != null)
                        {
                            cellValues.Add(row.GetCell(j).ToNotNullString());
                        }
                        else
                        {
                            cellValues.Add(string.Empty);
                        }
                    }

                    resultList.Add(buildResultItemFunc(cellValues));
                }
            }

            return resultList;
        }

        public static MemoryStream ToExcelStream(this IWorkbook book)
        {
            if (book.NumberOfSheets <= 0)
            {
                throw new Exception("无有效的sheet数据");
            }

            MemoryStream stream = new MemoryStream();

            stream.Seek(0, SeekOrigin.Begin);
            book.Write(stream);

            return stream;
        }

        public static byte[] ToExcelBytes(this IWorkbook book)
        {
            using (MemoryStream stream = ToExcelStream(book))
            {
                return stream.ToArray();
            }
        }

        public static JObject HttpUpload(this IWorkbook book, string uploadUrl, IDictionary<string, object> fieldData = null, string exportFileName = null)
        {
            using (HttpClient client = new HttpClient())
            {
                MultipartFormDataContent formData = new MultipartFormDataContent();
                ByteArrayContent fileContent = new ByteArrayContent(ToExcelBytes(book));
                //StreamContent fileContent = new StreamContent(ToExcelStream(book)); 二者都可以
                fileContent.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                fileContent.Headers.ContentDisposition = new ContentDispositionHeaderValue("form-data");

                if (string.IsNullOrWhiteSpace(exportFileName))
                {
                    exportFileName = Guid.NewGuid().ToString("N") + ((book is XSSFWorkbook) ? ".xlsx" : ".xls");
                }

                fileContent.Headers.ContentDisposition.FileName = exportFileName;
                fileContent.Headers.ContentDisposition.Name = "file";
                formData.Add(fileContent);

                Func<string, StringContent> getStringContent = (str) => new StringContent(str, Encoding.UTF8);

                if (fieldData != null)
                {
                    foreach (var header in fieldData)
                    {
                        formData.Add(getStringContent(header.Value.ToNotNullString()), header.Key);
                    }
                }


                HttpResponseMessage res = client.PostAsync(uploadUrl, formData).Result;
                string resContent = res.Content.ReadAsStringAsync().Result;
                return JObject.Parse(resContent);
            }
        }

        public static string SaveToFile(this IWorkbook book, string filePath)
        {
            if (File.Exists(filePath))
            {
                File.SetAttributes(filePath, FileAttributes.Normal);
                File.Delete(filePath);
            }

            using (FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                book.Write(fs);
            }

            return filePath;
        }


    }

}