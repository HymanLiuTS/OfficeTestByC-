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
using Excel = Microsoft.Office.Interop.Excel;

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
            Word.Application app = new Word.Application();
            Word.Document doc = app.Documents.Add("D:\\Test.docx");
            app=doc.Application;
            doc.ActiveWindow.Visible = true;
            foreach (Word.Bookmark bk in doc.Bookmarks)
            {
                if (bk.Name == "chart")
                {
                    /*方法二*/
                    /*启动Excel并填充数据*/
                    Excel.Application eApp = new Excel.Application();//创建Excel进程
                    eApp.Visible = true;//设置Excel可见
                    Excel.Workbook book = eApp.Workbooks.Add();//增加一个workboo
                    Excel.Worksheet sheet = eApp.Worksheets[1];//获取第一个Worksheet
                    Excel.Range range = sheet.get_Range("A1", "D2");//获取A1到D2范围内的Range
                    //向range中插入数据
                    range.Cells[1][1] = "姓名";
                    range.Cells[1][2] = "成绩";
                    range.Cells[2][1] = "张三";
                    range.Cells[2][2] = "89";
                    range.Cells[3][1] = "李四";
                    range.Cells[3][2] = "100";
                    range.Cells[4][1] = "王五";
                    range.Cells[4][2] = "95";
                    //插入图表
                    Excel.Chart xlChart = book.Charts.Add();
                    //设置图表源
                    xlChart.SetSourceData(range);
                    //拷贝表格
                    Word.Range wdRange = bk.Range;
                    range.Copy();
                    wdRange.Paste();
                   //拷贝图表数据到Word
                    wdRange.SetRange(wdRange.End, wdRange.End + 1);
                    xlChart.ChartArea.Copy();
                    wdRange.Paste();

                    /*方法三 只适用于office 2010及其以上版本*/
                   /* Word.Selection sel = app.Selection;
                    Word.InlineShape shape = sel.InlineShapes.AddChart();//插入图表
                    Word.Chart wdChart = shape.Chart;//获取图表
                    Word.ChartData chartData = wdChart.ChartData;//获取图表的数据
                    Excel.Workbook dataWorkbook = (Excel.Workbook)chartData.Workbook;//获取数据对应的workbook
                    dataWorkbook.Application.Visible = false;
                    Excel.Worksheet dataSheet = (Excel.Worksheet)dataWorkbook.Worksheets[1]; //获取图表对应的sheet
                    */
                    
                }
           }
            
            doc.SaveAs("E:\\Test.docx");
            app.Quit();
        }

    }
}
