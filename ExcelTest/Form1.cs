//////////////////////////////////////////////////////////////////////////
//////File: Form1.cs
//////Author: Hyman
//////Date: 2016/11/16
//////Description: 《C#中操作Excel（2）—— 新建、打开、保存和关闭Excel文档》源代码
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
            /* 以默认方式增加Excel */
            //app.Workbooks.Add)();
            /* 以模板方式增加Excel模板*/
            //app.Workbooks.Add("D:\\Test.xlsx");
            /* 打开一个Excel文档 */
            Workbook book=app.Workbooks.Open("D:\\Test.xlsx ");
            /* 增加一个sheet*/
            book.Worksheets.Add();
            /*保存*/
            book.Save();
            /*退出*/
            app.Quit();
        }

    }
}
