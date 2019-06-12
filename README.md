# ExcelEasyUtil
基于NPOI扩展封装的简易操作工具类库（简单灵活易用，支持导出、导入、上传等常见操作）

## 常见用法示例如下：
1. 第一种填充sheet方式：(重点在表格头的映射设置，通过Lamba属性表达式与要生成的EXCEL表头进行关联映射)

```c#
var book= ExcelEasyUtil.Core.CreateXlsxWorkBook()
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
               });
```

2. 第二种填充sheet方式：(重点在表格头类型设置【:XXX表示生成的EXCEL该列为某种格式的单元格，如：生日:DT表示是生日这列导出是日期类型格式】，第二个参数返回List<object>这个是可以很好的控制导出时需要的填充数据，可以自由控制数据，比如示例代码中额外增加了一列判断的数据列内容，第三个参数是为指定的列设置单元格的具体应用格式)

```c#
var book= ExcelEasyUtil.Core.CreateXlsxWorkBook()
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
               });
```

3. 第三种填充sheet方式：（重点仍然是在表头映射，由于这里的数据源是DataTable，故只是设置DataTable的列与EXCEL要导出的列名映射即可，无需指定列类型，第二个参数是为指定的列设置单元格的具体应用格式）

```c#
var book= ExcelEasyUtil.Core.CreateXlsxWorkBook() 
 .FillSheet("人员列表3", peoplesTable, //填充第三个表格
               new Dictionary<string, string> {
                   {"Name","姓名" },{"Birthday","生日" },{"Address","住址" },{"Education","学历" }, {"hasWork","有工作否" },{"Remark","备注" }
               }
               , new Dictionary<string, string>
               {
                   { "生日","yyyy-mm-dd"}
               });
```

### 由于实现了FillSheet方法后仍返回IWorkBook实例本身，即可采用链式的方式来快速完成多个sheet导出的，合并代码如下：
 
```c#
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
```

### 导入EXCEL数据（这里称为解析EXCEL数据）的示例用法：

```c#
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
```

完整说明详见作者博文：[【EXCEL终极总结分享】基于NPOI扩展封装的简易操作工具类库（简单灵活易用，支持导出、导入、上传等常见操作）](https://www.cnblogs.com/zuowj/p/10987224.html)

为了方便开发者使用，还封装成了NuGet包：　

> Packge Manager：Install-Package ExcelEasyUtil -Version 1.0.0

> .NET CLI：dotnet add package ExcelEasyUtil --version 1.0.0 

>  `<PackageReference Include="ExcelEasyUtil" Version="1.0.0" />`
