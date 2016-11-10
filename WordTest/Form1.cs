using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word=Microsoft.Office.Interop.Word;
using Graph=Microsoft.Office.Interop.Graph;

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
            Word.Document doc = app.Documents.Add("D:\\Test.docx");
            app=doc.Application;
            doc.ActiveWindow.Visible = true;
            foreach (Word.Bookmark bk in doc.Bookmarks)
            {
                if (bk.Name == "chart")
                {
                    object oClassType = "MSGraph.Chart.8";
                    Word.Range range = bk.Range;
                    Graph.Chart wdchart = (Graph.Chart)range.InlineShapes.AddOLEObject(oClassType).OLEFormat.Object;
                    wdchart.Application.DataSheet.Cells.Clear();//清空表格的初始数据
                    //axis.MaximumScale = 1;//最大刻度

                    //填充图表,起始的行号和列号都是1 
                    int i, j;
                    for (i = 0; i < 3; i++)//初始化列标头 
                    {
                        wdchart.Application.DataSheet.Cells[i + 1, 1] = "列" + i.ToString();
                    }
                    for (i = 0; i < 4; i++)//填充数据 
                    {
                        for (j = 0; j < 4; j++)
                        {
                            wdchart.Application.DataSheet.Cells[i + 2, j + 1] = i * j;
                        }
                    }

                    //根据Y轴来画图表 
                    wdchart.Application.PlotBy = Graph.XlRowCol.xlColumns;

                    wdchart.Legend.Delete();
                    wdchart.Height = 280;
                    wdchart.Width = 600;
                    //oShape.Height   =   oWord.InchesToPoints(3.57f);

                    //在图片之后添加文字
                    //range.InsertParagraphAfter();
                    //range.InsertAfter("sars");

                   /* Graph.Series ss = wdchart.SeriesCollection(1) as Graph.Series;
                    ss.Border.Color = 8388608;
                    ss.MarkerBackgroundColor = 8388608;
                    ss.MarkerForegroundColor = 8388608;*/

                    //更新图表并保存退出 
                    wdchart.Application.Update();
                    //wdchart.Application.Quit();
                   // wdchart = null;
                }
           }
            
            //doc.SaveAs("E:\\Test.docx");
            //app.Quit();
        }

    }
}
