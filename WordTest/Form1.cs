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
            Document doc = app.Documents.Add("D:\\Test.docx");
            doc.ActiveWindow.Visible = true;
            foreach (Bookmark bk in doc.Bookmarks)
            {
                bk.Range.Text = GetStrByBookmarkName(bk.Name);
           }
            doc.SaveAs("E:\\Test.docx");
            app.Quit();
        }

        private string GetStrByBookmarkName(string name)
        {
            string str = string.Empty;
            switch (name)
            {
                case "name":
                    str = "Hyman";
                    break;
                case "six":
                    str="男";
                    break;
                case "job":
                    str = "软件工程师";
                    break;
                case "date":
                    str = DateTime.Now.ToString();
                    break;
            }
            return str;
        }
    }
}
