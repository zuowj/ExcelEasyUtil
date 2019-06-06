using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ExcelEasyUtil
{
    /// <summary>
    /// 类型扩展方法集合
    /// author:zuowenjun 
    /// 2019-5-21
    /// </summary>
    public static class TypeExtension
    {
        /// <summary>
        /// 转换为不为空的字符串（即：若为空，则返回为空字符串，而不是Null）
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static string ToNotNullString(this object obj)
        {
            if (obj == null || obj == DBNull.Value)
            {
                return string.Empty;
            }
            return obj.ToString();
        }



        /// <summary>
        /// 判断列表中是否存在项
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public static bool HasItem(this IEnumerable<object> list)
        {
            if (list != null && list.Any())
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// 从字答串左边起取出指定长度的字符串
        /// </summary>
        /// <param name="str"></param>
        /// <param name="length"></param>
        /// <returns></returns>
        public static string Left(this string str, int length)
        {
            if (string.IsNullOrEmpty(str))
            {
                return string.Empty;
            }

            return str.Substring(0, length);
        }


        /// <summary>
        /// 从字答串右边起取出指定长度的字符串
        /// </summary>
        /// <param name="str"></param>
        /// <param name="length"></param>
        /// <returns></returns>
        public static string Right(this string str, int length)
        {
            if (string.IsNullOrEmpty(str))
            {
                return string.Empty;
            }

            return str.Substring(str.Length - length);
        }

        /// <summary>
        /// 判断DataSet指定表是否包含记录
        /// </summary>
        /// <param name="ds"></param>
        /// <param name="tableIndex"></param>
        /// <returns></returns>
        public static bool HasRows(this DataSet ds, int tableIndex = 0)
        {
            if (ds != null && ds.Tables[tableIndex].Rows.Count > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// 通用类型转换方法，EG:"".As<String>()
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static T As<T>(this object obj)
        {
            T result;
            try
            {
                Type type = typeof(T);
                if (type.IsNullableType())
                {
                    if (obj == null || obj.ToString().Length == 0)
                    {
                        result = default(T);
                    }
                    else
                    {
                        type = type.GetGenericArguments()[0];
                        result = (T)Convert.ChangeType(obj, type);
                    }
                }
                else
                {
                    if (obj == null)
                    {
                        if (type == typeof(string))
                        {
                            result = (T)Convert.ChangeType(string.Empty, type);
                        }
                        else
                        {
                            result = default(T);
                        }
                    }
                    else
                    {
                        result = (T)Convert.ChangeType(obj, type);
                    }
                }
            }
            catch
            {
                result = default(T);
            }

            return result;
        }

        /// <summary>
        /// 判断是否为可空类型
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static bool IsNullableType(this Type type)
        {
            return (type.IsGenericType &&
              type.GetGenericTypeDefinition().Equals
              (typeof(Nullable<>)));
        }


        /// <summary>
        /// 将为DateTime的Object类型转换成指定日期格式的字符串
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="hasYMD">是否包含年月日</param>
        /// <param name="hasHM">是否包含时分</param>
        /// <param name="hasS">是否包含秒</param>
        /// <returns></returns>
        public static string ToFormatDateTimeString(this object obj, bool hasYMD = true, bool hasHM = true, bool hasS = true)
        {
            string formatStr = string.Empty;
            if (hasYMD)
            {
                formatStr = "yyyy-MM-dd";
            }

            if (hasHM)
            {
                formatStr += (hasYMD ? " " : "") + "HH:mm";
            }

            if (hasS)
            {
                formatStr += (hasHM ? ":" : "") + "ss";
            }

            DateTime sqlMinDt = DateTime.Parse("1753-01-01");//SQL 最小日期时间

            if (string.IsNullOrWhiteSpace(obj.ToString()) || obj == DBNull.Value)
            {
                obj = sqlMinDt;
            }

            DateTime dtValue = Convert.ToDateTime(obj);
            if (dtValue < sqlMinDt)
            {
                dtValue = sqlMinDt;
            }

            return dtValue.ToString(formatStr);
        }

        /// <summary>
        /// 字符串是否有值
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static bool HasValue(this string str)
        {
            return !string.IsNullOrWhiteSpace(str);
        }
    }
}