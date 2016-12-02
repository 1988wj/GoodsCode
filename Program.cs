using System;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace _1988wj.GoodsCode
{
    //2016-12-01编辑整理
    class Program
    {
        /// <summary>
        /// 主函数
        /// </summary>
        /// <param name="args">拖拽文件的文件路径</param>
        static void Main(string[] args)
        {
            Console.WriteLine("===================================软件说明===================================");
            Console.WriteLine("  Excel格式对账单转换为开票软件能导入的txt格式 商品编码 ");
            Console.WriteLine("  自动模糊匹配Excel文件中包含以下关键字的列:\r\n");
            Console.WriteLine("          '名称' =>   用作 [名称]     列");
            Console.WriteLine("          '颜色' => 附加到 [名称]     列末尾");
            Console.WriteLine("          '单位' =>   用作 [计量单位] 列");
            Console.WriteLine("          '单价' =>   用作 [单价]     列");
            Console.WriteLine("  '款号'或'编号' => 附加到 [名称]     列开头");
            Console.WriteLine("  '规格'或'型号' =>   用作 [规格型号] 列\r\n");
            Console.WriteLine("  简码, 商品税目, 税率, 含税价标志, 隐藏标志, 中外合作油气田, 税收分类编码");
            Console.WriteLine("  是否享受优惠政策, 税收分类编码名称, 优惠政策类型, 零税率标识, 编码版本号\r\n");
            Console.WriteLine("  以上字段可在 [{0}.config] 中设置默认填充值", AppDomain.CurrentDomain.SetupInformation.ApplicationName);
            Console.WriteLine("==============================================================================");
            //(支持多文件)Excel文件拖拽至程序上直接编码
            foreach (var item in args)
            {
                Run(item);
            }
            //输入文件路径执行编码
            string filePath;
            while (true)
            {
                Console.WriteLine("\r\n请输入Excel文件地址(直接回车退出):");
                filePath = Console.ReadLine();
                filePath = filePath.ToLower().Trim();
                if (filePath.Length == 0)
                {
                    return;
                }
                else
                {
                    Run(filePath);
                }
            }
        }
        /// <summary>
        /// 对文件编码及保存
        /// </summary>
        /// <param name="excelPath">Excel文件路径</param>
        static void Run(string excelPath)
        {
            string fileName;
            string savePath;

            if (File.Exists(excelPath))
            {
                //转为绝对路径
                excelPath = Path.GetFullPath(excelPath);
                fileName = Path.GetFileNameWithoutExtension(excelPath);
                fileName = Regex.Match(fileName, @"\D+").Value.Replace("对账单", "").Trim();
                //生成商品编码文件路径
                savePath = string.Format("{0}\\Goods[{1}]{2}.txt", Path.GetDirectoryName(excelPath), fileName, DateTime.Now.ToString("yyMMdd"));
                Console.WriteLine("正在对文件\"{0}\"进行编码...", excelPath);
                try
                {
                    string codeString = MyCoding(excelPath).ToString();
                    File.WriteAllText(savePath, codeString, Encoding.Default);
                    Console.WriteLine("已保存到\"{0}\"\r\n", savePath);
                }
                catch (InvalidOperationException)
                {
                    Console.WriteLine("Microsoft.Jet.OLEDB.4.0 只支持32位程序");
                    Console.WriteLine("如果需要运行64位版本 需安装AccessDatabaseEngine提供访问支持");
                }
                catch (OleDbException)
                {
                    Console.WriteLine("文件无法正常打开!");
                    Console.WriteLine("如果Excel文件格式为 xlsx 需安装AccessDatabaseEngine提供访问支持");
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
            else
            {
                Console.WriteLine("文件地址:\"{0}\"错误!", excelPath);
            }
        }
        /// <summary>
        /// 生成商品编码
        /// </summary>
        /// <param name="excelPath">Excel文件路径</param>
        /// <returns>商品编码</returns>
        static StringBuilder MyCoding(string excelPath)
        {
            int 名称列 = -1;
            int 单位列 = -1;
            int 单价列 = -1;
            int 颜色列 = -1;
            int 编号列 = -1;
            int 规格列 = -1;
            string 编号;
            string 颜色;
            string 编码 = ConfigurationManager.AppSettings["编码"];
            string 名称 = ConfigurationManager.AppSettings["名称"];
            string 简码 = ConfigurationManager.AppSettings["简码"];
            string 税目 = ConfigurationManager.AppSettings["商品税目"];
            string 税率 = ConfigurationManager.AppSettings["税率"];
            string 规格 = ConfigurationManager.AppSettings["规格型号"];
            string 单位 = ConfigurationManager.AppSettings["计量单位"];
            string 单价 = ConfigurationManager.AppSettings["单价"];
            string 含税 = ConfigurationManager.AppSettings["含税价标志"];
            string 隐藏 = ConfigurationManager.AppSettings["隐藏标志"];
            string 油气 = ConfigurationManager.AppSettings["中外合作油气田"];
            string 税码 = ConfigurationManager.AppSettings["税收分类编码"];
            string 优惠 = ConfigurationManager.AppSettings["是否享受优惠政策"];
            string 税名 = ConfigurationManager.AppSettings["税收分类编码名称"];
            string 政策 = ConfigurationManager.AppSettings["优惠政策类型"];
            string 零税 = ConfigurationManager.AppSettings["零税率标识"];
            string 版本 = ConfigurationManager.AppSettings["编码版本号"];
            StringBuilder codeString = new StringBuilder(
                "{商品编码}[分隔符]\"~~\"\r\n" +
                "// 每行格式 :\r\n" +
                "// 编码~~名称~~简码~~商品税目~~税率~~规格型号~~计量单位~~单价~~含税价标志~~隐藏标志~~中外合作油气田~~税收分类编码~~是否享受优惠政策~~税收分类编码名称~~优惠政策类型~~零税率标识~~编码版本号\r\n",
                1024);
            int lineNumber;
            int columnCount;
            string tableName;
            string columnName;
            string cmdString;
            DataTable tableSchema;
            OleDbCommand cmd;
            OleDbConnection conn;
            OleDbDataReader dataReader;

            //用数据库访问方式读取Excel文档
            try
            {   //支持xlsx需要安装AccessDatabaseEngine
                //HDR= Yes:首行为标题 No:行为数据 IMEX= 0:只写 1:只读 2:读写
                conn = new OleDbConnection(string.Format(
                    "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='{0}';Extended Properties='Excel 12.0; HDR=Yes; IMEX=1'", excelPath));
                conn.Open();
            }
            catch (InvalidOperationException)
            {   //只能编译为 32位 程序才能通过 Microsoft.Jet.OLEDB.4.0 访问xls文件
                conn = new OleDbConnection(string.Format(
                    "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='{0}';Extended Properties='Excel 8.0; HDR=Yes; IMEX=1'", excelPath));
                conn.Open();
            }
            //获取表名
            tableSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            tableName = tableSchema.Rows[0]["Table_Name"].ToString();
            Console.WriteLine("找到工作表[{0}]", tableName);
            //生成查询字符串
            cmdString = string.Format("select * from [{0}]", tableName);
            Console.WriteLine("生成查询命令=> {0}", cmdString);
            //查询数据
            cmd = new OleDbCommand(cmdString, conn);
            dataReader = cmd.ExecuteReader();
            Console.WriteLine("执行查询操作=>");
            //获取数据表列数
            columnCount = dataReader.FieldCount;
            Console.WriteLine("共发现{0}列数据:", columnCount);
            //匹配数据列
            for (int i = 0; i < columnCount; i++)
            {
                columnName = dataReader.GetName(i);
                Console.Write("{0}.'{1}'", i, columnName);

                if (名称列 == -1 && (columnName.Contains("名称"))) { 名称列 = i; Console.WriteLine("=>用作 [名称] 列"); continue; }

                if (单位列 == -1 && (columnName.Contains("单位"))) { 单位列 = i; Console.WriteLine("=>用作 [计量单位] 列"); continue; }

                if (单价列 == -1 && (columnName.Contains("单价"))) { 单价列 = i; Console.WriteLine("=>用作 [单价] 列"); continue; }

                if (颜色列 == -1 && (columnName.Contains("颜色"))) { 颜色列 = i; Console.WriteLine("=>附加到 [名称] 列末尾"); continue; }

                if (编号列 == -1 && (columnName.Contains("款号") || columnName.Contains("编号"))) { 编号列 = i; Console.WriteLine("=>附加到 [名称] 列开头"); continue; }

                if (规格列 == -1 && (columnName.Contains("规格") || columnName.Contains("型号"))) { 规格列 = i; Console.WriteLine("=>用作 [规格型号] 列"); continue; }

                Console.WriteLine("=>未使用");
            }
            //未匹配到合适的[名称],[计量单位],[单价]列抛出异常
            if (名称列 == -1) { throw new ArgumentNullException("[名称]", "未匹配到合适的[名称]列"); }

            if (单位列 == -1) { throw new ArgumentNullException("[计量单位]", "未匹配到合适的[计量单位]列"); }

            if (单价列 == -1) { throw new ArgumentNullException("[单价]", "未匹配到合适的[单价]列"); }
            //生成并保存商品编码
            Console.WriteLine("开始生成商品编码");

            //string nameString, specificationsString, unitString, priceString;

            lineNumber = 0;
            while (dataReader.Read())
            {
                //[名称]
                名称 = dataReader[名称列].ToString().Trim();
                if (名称.Length == 0) { continue; }

                if (编号列 != -1) { 编号 = dataReader[编号列].ToString().Trim() + " "; }
                else { 编号 = ""; }

                if (颜色列 != -1) { 颜色 = " " + dataReader[颜色列].ToString().Trim(); }
                else { 颜色 = ""; }

                //[名称]整合编号和颜色并添加序号
                lineNumber++;
                名称 = string.Format("{0}{1}{2}({3})", 编号, 名称, 颜色, lineNumber);

                //编码
                编码 = lineNumber.ToString("000");

                //[规格型号]
                if (规格列 != -1)
                {
                    规格 = dataReader[规格列].ToString().Trim();
                }
                else
                {
                    规格 = "";
                }

                //[计量单位]
                单位 = dataReader[单位列].ToString().Trim();
                if (单位 == null || 单位.Length == 0)
                {
                    单位 = ConfigurationManager.AppSettings["计量单位"];
                }

                //[单价]
                单价 = dataReader[单价列].ToString().Trim();
                if (double.TryParse(单价, out double result))
                {
                    单价 = "0";
                }

                //生成行信息
                codeString.AppendFormat("{0}~~{1}~~{2}~~{3}~~{4}~~{5}~~{6}~~{7}~~{8}~~{9}~~{10}~~{11}~~{12}~~{13}~~{14}~~{15}~~{16}\r\n",
                    编码, 名称, 简码, 税目, 税率, 规格, 单位, 单价, 含税, 隐藏, 油气, 税码, 优惠, 税名, 政策, 零税, 版本);
            }
            conn.Close();
            Console.WriteLine("共生成[ {0} ]条信息", lineNumber);
            return codeString;
        }
    }
}