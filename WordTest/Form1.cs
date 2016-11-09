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
                if (bk.Name == "name")
                {
                    bk.Range.Text = "Hyman";
                }
                else if (bk.Name == "picture")
                {
                    bk.Select();
                    Selection sel = app.Selection;
                    sel.InlineShapes.AddPicture("D:\\Test.jpg");
                }
               
           }
            
            doc.SaveAs("E:\\Test.docx");
            app.Quit();
        }

    }
}
