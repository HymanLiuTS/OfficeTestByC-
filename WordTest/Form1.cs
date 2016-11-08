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
            string version = app.Version;
            Console.WriteLine(version);
            Document doc = app.Documents.Open("D:\\Test.docx");
            doc.ActiveWindow.Visible = true;
            Range range = doc.Range();
            range.Font.Size = 14;
            range.Font.Name = "微软雅黑";
            range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            range.Bold = 10;
            range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineDotDash;
            range.Select();
            doc.Save();
            //app.Quit();
            app = null;

        }
    }
}
