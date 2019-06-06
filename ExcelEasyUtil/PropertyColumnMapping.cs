using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Text;

namespace ExcelEasyUtil
{
    /// <summary>
    /// 类属性与表格列映射类
    /// author:zuowenjun
    /// 2019-5-30
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class PropertyColumnMapping<T> : Dictionary<string, Expression<Func<T, dynamic>>>
    {
        public PropertyColumnMapping<T> Append(string columnName, Expression<Func<T, dynamic>> selectProperty)
        {
            Add(columnName, selectProperty);
            return this;
        }
    }
}
