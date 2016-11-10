using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace WordTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Document doc = app.Documents.Add("D:\\Test.docx");
            app=doc.Application;
            doc.ActiveWindow.Visible = true;
            foreach (Bookmark bk in doc.Bookmarks)
            {
                if (bk.Name == "marks")
                {
                    Range range = bk.Range;
                    range.Tables.Add(range,3,2);
                    Table tb = range.Tables[1];
                    tb.set_Style("网格型");
                    tb.Cell(1, 1).Range.Text = "姓名";
                    tb.Cell(1, 2).Range.Text = "成绩";
                    tb.Cell(2, 1).Range.Text = "张三";
                    tb.Cell(2, 2).Range.Text = "89";
                    tb.Cell(3, 1).Range.Text = "李四";
                    tb.Cell(3, 2).Range.Text = "98";  
                }
           }
            
            doc.SaveAs("E:\\Test.docx");
            app.Quit();
        }

    }
}
