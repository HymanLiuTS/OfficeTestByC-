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
            Console.WriteLine(range.Text.ToString());
            range = doc.Range(4, 10);
            range.Text = "新?的Ì?测a试º?文?档Ì¦Ì。¡ê。¡ê";
            doc.Save();
            app.Quit();
            app = null;

        }
    }
}
