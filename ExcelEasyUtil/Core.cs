using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelEasyUtil
{
    /// <summary>
    /// NPOI 相关核心入口方法帮助类
    /// author:zuowenjun
    /// 2019-5-21
    /// </summary>
    public static class Core
    {
        /// <summary>
        /// 创建一个基本XLSX格式的EXCEL工作薄对象
        /// </summary>
        /// <returns></returns>
        public static IWorkbook CreateXlsxWorkBook()
        {
            return new XSSFWorkbook();
        }

        /// <summary>
        /// 创建一个基本XLS格式的EXCEL工作薄对象
        /// </summary>
        /// <returns></returns>
        public static IWorkbook CreateXlsWorkBook()
        {
            return new HSSFWorkbook();
        }

        /// <summary>
        /// 打开指定文件的EXCEL工作薄对象
        /// </summary>
        /// <param name="filePath">Excel文件路径</param>
        /// <returns></returns>
        public static IWorkbook OpenWorkbook(string filePath)
        {
            bool isCompatible = filePath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase);
            var fileStream = System.IO.File.OpenRead(filePath);
            if (isCompatible)
            {
                return new HSSFWorkbook(fileStream);
            }
            else
            {
                return new XSSFWorkbook(fileStream);
            }
        }

    }
}
