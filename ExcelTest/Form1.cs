//////////////////////////////////////////////////////////////////////////
//////File: Form1.cs
//////Author: Hyman
//////Date: 2016/11/23
//////Description: C#中操作Excel（5）—— 获取Excel中数据的两种方法
//////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Reflection;
using System.Data.OleDb;

namespace ExcelTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;
            /* 以模板方式打开Excel模板*/
            Workbook book = app.Workbooks.Add("D:\\Test.xlsx");
            /* 获取第一个sheet*/
            Worksheet sheet = book.Worksheets[1];
            
            Range range = sheet.get_Range("A1");
            range = range.get_Resize(4, 2);
           
            /*获取Excel中的数据*/
            object _optionalValue = Missing.Value;
            object[,] objects = new object[4, 2];
            objects = range.get_Value(_optionalValue);

            /*插入数据*/
            Range rangeToInsert = sheet.get_Range("D1","E4");
            rangeToInsert.set_Value(_optionalValue,objects);
           
            /*保存*/
            //book.SaveAs("E:\\Test.xlsx ");
            /*退出*/
            //app.Quit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            /*定义连接字符串*/
            bool hasTitle = false;
            string path="D:\\Test.xlsx";
            string fileType = System.IO.Path.GetExtension(path);
            string strCon = string.Empty;
            if (fileType == ".xls")
            {
                strCon = string.Format("Provider=Microsoft.Jet.OLEDB.{0}.0;" +
                            "Extended Properties=\"Excel {1}.0;HDR={2};IMEX=1;\";" +
                            "Data Source={3};",
                            (fileType == ".xls" ? 4 : 12), (fileType == ".xls" ? 8 : 12), (hasTitle ? "Yes" : "NO"), path);
            }
            else
            {
                strCon = string.Format("Provider=Microsoft.ACE.OLEDB.{0}.0;" +
                            "Extended Properties=\"Excel {1}.0;HDR={2};IMEX=1;\";" +
                            "Data Source={3};",
                            (fileType == ".xls" ? 4 : 12), (fileType == ".xls" ? 8 : 12), (hasTitle ? "Yes" : "NO"), path);
            }

            /*使用SQL语句读取数据*/
            string sheetName = "sheet1";
            string strCom = " SELECT * FROM [" + sheetName + "$]";

            /*建立数据库连接*/
            OleDbConnection myConn = new OleDbConnection(strCon);

            /*建立sql语句执行器*/
            OleDbDataAdapter myCommand = new OleDbDataAdapter(strCom, myConn);

            /*执行并读取数据*/
            System.Data.DataTable dt = new System.Data.DataTable();
            myConn.Open();//打开连接
            myCommand.Fill(dt);//填充数据
            myConn.Close();//关闭连接

            /*将数据集里的数据写入到Excel*/
            //1.创建连接字符串

            String sConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" +

     "Data Source=D:/Test2.xls;" +

     "Extended Properties=Excel 8.0;";

            OleDbConnection cn = new OleDbConnection(sConnectionString);

            string sqlCreate = "CREATE TABLE TestSheet ([姓名] VarChar,[成绩] INTEGER)";

            OleDbCommand cmd = new OleDbCommand(sqlCreate, cn);

            //创建Excel文件：D:/Test2.xls

            cn.Open();

            //创建TestSheet工作表

            cmd.ExecuteNonQuery();

            //添加数据
            for (int i = 1; i < dt.Rows.Count;i++ )
            {
                DataRow row = dt.Rows[i];
                cmd.CommandText = "INSERT INTO TestSheet VALUES('" + row["F1"] + "'," + row["F2"] + ")";
                cmd.ExecuteNonQuery();
            }

            //关闭连接
            cn.Close();
        }

    }
}
