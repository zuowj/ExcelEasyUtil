using System;
using System.Collections.Generic;
using ExcelEasyUtil;
using System.Linq.Expressions;
using System.IO;
using System.Data;

namespace Examples
{
    class Program
    {
        static void Main(string[] args)
        {
            var peoples = BuildTestDataModelList();
            var peoplesTable = BuildTestDataTable();

            string savedPath = ExcelEasyUtil.Core.CreateXlsxWorkBook()
                .FillSheet("人员列表1", peoples,//填充第一个表格
                                            //new Dictionary<string, Expression<Func<People, dynamic>>> //设置表格头，原始类型
               new PropertyColumnMapping<People> //设置表格头，专用简化类型
               {
                {"姓名",p=>p.Name },{"年龄",p=>p.Age },{"生日",p=>p.Birthday },{"住址",p=>p.Address },{"学历",p=>p.Education },
                { "有工作否",p=>p.hasWork },{"备注",p=>p.Remark }
               },
               new Dictionary<string, string> //为指定的列设置单元格格式
               {
                { "年龄","0岁"},{"生日","yyyy年mm月dd日"}
               })
               .FillSheet("人员列表2", peoples,  //填充第二个表格
               new List<string>
               {
                   "姓名","年龄:NUM","生日:DT","住址","学历","有工作否:BL","备注","额外填充列"
               }, (p) =>
               {
                   return new List<object> {
                       p.Name,p.Age,p.Birthday,p.Address,p.Education,p.hasWork?"有":"无",p.Remark,(p.Age<=30 && p.hasWork)?"年轻有为":"要么老了要么没工作，生活堪忧"
                   };
               }, new Dictionary<string, string>
               {
                   { "生日","yyyy-mm-dd"}
               })
               .FillSheet("人员列表3", peoplesTable, //填充第三个表格
               new Dictionary<string, string> {
                   {"Name","姓名" },{"Birthday","生日" },{"Address","住址" },{"Education","学历" }, {"hasWork","有工作否" },{"Remark","备注" }
               }
               , new Dictionary<string, string>
               {
                   { "生日","yyyy-mm-dd"}
               })
               .SaveToFile(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "testdata123.xlsx"));

            Console.WriteLine("导出EXCEL文件路径：" + savedPath);

            var xlsTable = ExcelEasyUtil.Core.OpenWorkbook(savedPath).ResolveDataTable("人员列表1", 0);
            foreach (DataRow row in xlsTable.Rows)
            {
                string rowStr = string.Join("\t", row.ItemArray);
                Console.WriteLine(rowStr);
            }

            var xlsPeoples = ExcelEasyUtil.Core.OpenWorkbook(savedPath).ResolveAs<People>("人员列表2", 0, list =>
            {
                return new People
                {
                    Name = list[0],
                    Birthday = ConvertToDate(list[2]),//日期处理相对较麻烦
                    Address = list[3]
                };
            }, 0, 4);

            Console.WriteLine("-".PadRight(30,'-'));
            foreach (var p in xlsPeoples)
            {
                string rowStr = string.Join("\t", p.Name, p.Age, p.Birthday, p.Address);
                Console.WriteLine(rowStr);
            }

            Console.ReadLine();
        }

        private static List<People> BuildTestDataModelList()
        {
            var peoples = new List<People>();
            DateTime tmpDate = DateTime.Today;
            for (int i = 1; i <= 100; i++)
            {
                tmpDate = tmpDate.AddMonths(-6);
                peoples.Add(new People
                {
                    Name = "测试人-" + i.ToString(),
                    Address = $"中国深圳XX区XX路第{i}号",
                    Birthday = tmpDate,
                    Education = (i % 2 == 0) ? "大学" : "小学",
                    hasWork = (i % 5 == 0),
                    Remark = "测试数据-" + i.ToString()
                });
            }

            return peoples;
        }

        public static DataTable BuildTestDataTable()
        {
            var table = new DataTable();
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Address", typeof(string));
            table.Columns.Add("Birthday", typeof(DateTime));
            table.Columns.Add("Education", typeof(string));
            table.Columns.Add("hasWork", typeof(bool));
            table.Columns.Add("Remark", typeof(string));

            DateTime tmpDate = DateTime.Today;
            for (int i = 1; i <= 100; i++)
            {
                tmpDate = tmpDate.AddMonths(-6);
                table.Rows.Add("测试人-" + i.ToString(),
                    $"中国深圳XX区XX路第{i}号",
                    tmpDate,
                    (i % 2 == 0) ? "大学" : "小学",
                    (i % 5 == 0),
                    "测试数据-" + i.ToString());
            }

            return table;
        }


        /// <summary>
        /// 转化日期
        /// </summary>
        /// <param name="date">日期</param>
        /// <returns></returns>
        public static DateTime ConvertToDate(object date)
        {
            try
            {
                return Convert.ToDateTime(date);
            }
            catch { }

            string dtStr = (date ?? "").ToString();

            DateTime dt = new DateTime();

            if (DateTime.TryParse(dtStr, out dt))
            {
                return dt;
            }

            try
            {
                string spStr = "";
                if (dtStr.Contains("-"))
                {
                    spStr = "-";
                }
                else if (dtStr.Contains("/"))
                {
                    spStr = "/";
                }

                dtStr = System.Text.RegularExpressions.Regex.Replace(dtStr, "[年月日]+", string.Empty);

                string[] time = dtStr.Split(spStr.ToCharArray());
                int year = Convert.ToInt32(time[2]);
                int month = Convert.ToInt32(time[0]);
                int day = Convert.ToInt32(time[1]);
                string years = Convert.ToString(year);
                string months = Convert.ToString(month);
                string days = Convert.ToString(day);
                if (months.Length == 4)
                {
                    dt = Convert.ToDateTime(date);
                }
                else
                {
                    string rq = years + "-" + months + "-" + days;
                    dt = Convert.ToDateTime(rq);
                }
            }
            catch
            {
                throw new Exception("日期格式不正确，转换日期类型失败！");
            }
            return dt;
        }

    }
}
