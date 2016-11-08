using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

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
            app.Visible = true;
            //1新建操作
            //1.1 按照默认方式新建文档
            //app.Documents.Add();
            //1.2 按照自定义模板创建文档
            //app.Documents.Add("D://Test.docx");
            //2 打开文档
            app.Documents.Open("D://Test.docx");
            //3 保存文档
            app.Documents.Save();
            //4 退出word
            app.Quit();
        }
    }
}
