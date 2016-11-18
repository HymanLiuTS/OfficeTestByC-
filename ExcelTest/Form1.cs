//////////////////////////////////////////////////////////////////////////
//////File: Form1.cs
//////Author: Hyman
//////Date: 2016/11/18
//////Description: 《 C#中操作Excel（4）—— 向Excel中插入图表的两种方式》源代码
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
            /*获取Range对象*/
            Range range = sheet.get_Range("B2", "C4");
            /*准备数据源*/
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("Name");
            dt.Columns.Add("Marks");
            DataRow row1 = dt.NewRow();
            row1["Name"] = "Tom";
            row1["Marks"] = "96";
            dt.Rows.Add(row1);
            DataRow row2 = dt.NewRow();
            row2["Name"] = "Jerry";
            row2["Marks"] = "91";
            dt.Rows.Add(row2);
            DataRow row3 = dt.NewRow();
            row3["Name"] = "Pooly";
            row3["Marks"] = "100";
            dt.Rows.Add(row3);
            /*插入数据*/
            for (int i = 0; i < dt.Rows.Count;i++ )
            {
                for (int j = 0; j < dt.Columns.Count;j++ )
                {
                    
                    range[j+1][i+1] = dt.Rows[i][j];
                }
            }
            /*插入和sheet同级别的chart*/
            Chart chart = book.Charts.Add();
            /*设置表格的数据源*/
            chart.SetSourceData(range);

            /*插入sheet的子图表*/
            Microsoft.Office.Interop.Excel.Shape shape = sheet.Shapes.AddChart();
            /*设置表格的数据源*/
            chart = shape.Chart;
            chart.SetSourceData(range);

         

           
            /*保存*/
            //book.SaveAs("E:\\Test.xlsx ");
            /*退出*/
            //app.Quit();
        }

    }
}
