using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using Excel = Microsoft.Office.Interop.Excel;

namespace 批量验证程序V1._0
{
    public partial class Form1 : Form
    {
        /// <summary>
        /// 玩法设定ssc=5，sc=10
        /// </summary>
        int LotCount;
        //object[,] ExcelData;
        Calc exCalc = new Calc();


        /// <summary>
        /// 主pwws
        /// </summary>
        public Form1()
        {
            InitializeComponent();
            //string fpath = string.Format(@"E:\制作\极速时时彩分析\测试文件目录\btExporttest{0:HHmmss}.xlsx", DateTime.Now);//设定导出文件名及路径;
            //exCalc.app.Visible = true;
            //Excel.Workbook testwk = exCalc.CreateWk(fpath,10);
            //foreach (Worksheet st in testwk.Worksheets)
            //{
            //    MessageBox.Show(st.Name);
            //}
            LotCount = JSSC.Checked == true ? 10 : 5;

            AddIitemsForLBAndCB();
        }


        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            //DateTime dt;
            //dt = dateTimePicker1.Value;
            //DateTime dt1;
            //dt1 = dt.AddDays(1);
            //MessageBox.Show($"dt:{dt},dt1:{dt1}");
        }


        /// <summary>
        /// 极速赛车
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            LotCount = JSSC.Checked == true ? 10 : 5;
            AddIitemsForLBAndCB();

        }



        /// <summary>
        /// 给ListBox及comboBox添加Items
        /// </summary>
        public void AddIitemsForLBAndCB()
        {
            this.listBox1.Items.Clear();
            this.comboBox1.Items.Clear();
            for (int i = 0; i < LotCount; i++)
            {
                this.listBox1.Items.Add(string.Format("第{0:D2}赛道", i+1));
                //if (i < LotCount - 1) { this.comboBox1.Items.Add((i + 1).ToString()); }
            }
            for (int i = 0; i <9; i++)
            {
                this.comboBox1.Items.Add((i + 1).ToString());
            }
        }



        /// <summary>
        /// 极速时时彩
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {

        }



        /// <summary>
        /// 测试周期数
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            textBox1.Text = GetNum(textBox1.Text);
        }



        /// <summary>
        /// 验证textBox是否为数字
        /// </summary>
        /// <param name="tText">textBox的值</param>
        /// <returns></returns>
          public string GetNum(string tText)
        {
            Regex rx = new Regex("^[0-9]*$");
            if (rx.IsMatch(tText))
            {
                return tText;
            }
            else
            {
                MessageBox.Show("输入值错误,请输入数字！");
                this.textBox1.TextChanged -= new EventHandler(textBox1_TextChanged);
                textBox1.Clear();
                this.textBox1.TextChanged += new EventHandler(textBox1_TextChanged);
                textBox1.Focus();
                return textBox1.Text;
            }
        }



        /// <summary>
        /// 读取csv文件数据
        /// </summary>
        /// <returns>返回一个List数组</returns>
        public List<string> GetFiles()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            //设置对话框的标题
            ofd.Title = "请选择数据文件";
            //设置对话框可以多选
            ofd.Multiselect = true;
            //设置对话框的初始目录
            ofd.InitialDirectory = System.IO.Directory.GetCurrentDirectory();
            //设置对话框的文件类型
            ofd.Filter = "CSV文件(*.csv)|*.csv";
            ////展示对话框
            //ofd.ShowDialog();

            //获得在打开对话框中选中文件的路径

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string path = ofd.FileName;
                var encoding = Encoding.GetEncoding("gbk");
                var lines = new List<string>(File.ReadLines(path, encoding));
                var line1 = new List<string>();
                /*                for (int i = lines.Count; i==0 ; i--)
                                {
                                    if (lines[i]==",,")
                                    {
                                        lines.Remove(lines[i]);
                                    }
                                }*/
                for (int i = 0; i < lines.Count; i++)
                {
                    if (lines[i] != ",,")
                    {
                        line1.Add(lines[i]);
                    }
                    else
                    {
                        break;
                    }
                }
                return line1;
            }
            else
            {
                return null;

            }

        }



/*        /// <summary>
        /// 调用Com接口操作Excel，读取数据
        /// </summary>
        /// <param name="exFileName">csv路径</param>
        /// <returns></returns>
        public void ReadExcel(string exFileName)
        {
            Excel.Application app =new Excel.Application();
            Workbook wk = app.Workbooks.Open(exFileName);
            Worksheet st = wk.Sheets[1];
            ExcelData = (object[,])st.UsedRange.Value;
        }
*/


        /// <summary>
        /// 打开Excel文件并读取数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 打开文件ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            //设置对话框的标题
            ofd.Title = "请选择要打开的数据文件";
            //设置对话框可以多选
            ofd.Multiselect = false;
            //设置对话框的初始目录
            ofd.InitialDirectory = System.IO.Directory.GetCurrentDirectory();
            //设置对话框的文件类型
            ofd.Filter = "CSV文件(*.csv)|*.csv";
            while (true)
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    exCalc.ExFilePath = ofd.FileName;
                    exCalc.ReadExcel(ofd.FileName);
                    break;
                }
            }

        }



        /// <summary>
        /// 主程序
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            /// 判断参数是否已经设置
            if (string.IsNullOrEmpty(comboBox1.Text))
            {
                MessageBox.Show("选择过滤基数");
                return;
            }
            else if (listBox1.SelectedItems.Count == 0)
            {
                MessageBox.Show("请选择测试赛道");
                return;
            }
            else if (string.IsNullOrEmpty(textBox1.Text))
            {
                MessageBox.Show("请输入测试周期");
                return;
            }

            exCalc.LotCount = LotCount;//开奖号位数

            exCalc.FilterType =MaxCheck.Checked ? "Max" : "Min";//过滤方式的确定。


            int MaxNum = JSSC.Checked ? 10 : 9;
            exCalc.MaxNum = MaxNum;

            this.comboBox1.Text = this.comboBox1.Text;
            this.textBox1.Text = this.textBox1.Text;

            exCalc.FilterCount = int.Parse(this.comboBox1.Text);


            int[] temparr=new int[listBox1.SelectedIndices.Count];
            listBox1.SelectedIndices.CopyTo(temparr, 0);//选择的赛道，因为从0开始因此需要加1
            exCalc.LotSD = temparr;

            dateTimePicker1.Value = new DateTime(2020, 3, 1, 0, 0,0);
            dateTimePicker2.Value = new DateTime(2020, 3, 2, 23, 59,0);

            dateTimePicker3.Value = new DateTime(2020, 4, 1, 0, 0, 0);
            dateTimePicker4.Value = new DateTime(2020, 4, 5, 23, 59, 0);

            
            

            ///把基础数据导出
            exCalc.GetBaseFilterData(SingleDayCheck.Checked, dateTimePicker1.Value, dateTimePicker2.Value, dateTimePicker3.Value, dateTimePicker4.Value,
                 int.Parse(this.comboBox1.Text), int.Parse(this.textBox1.Text));


            ///导出测试数据
            ///测试数据由基础数据对应生成。因为在类中需要同步计算。


            //if (this.SingleDayCheck.Checked)
            //{

            //}
            //else
            //{
            //    int[][,] BaseTen = exCalc.GetTenFive(exCalc.FilterData(dateTimePicker1.Value, dateTimePicker2.Value));//多日的基础数据是固定的
            //    //因此不用循环

            //    string fpath = string.Format(@"E:\制作\极速时时彩分析\测试文件目录\btExporttest{0:HHmmss}.xlsx", DateTime.Now);//设定导出文件名及路径;

            //    exCalc.wkten = exCalc.CreateWk(fpath);
            //    for (int i = 0; i < BaseTen.GetLength(0); i++)
            //    {
            //        string bt = this.listBox1.Items[exCalc.LotSD[i]].ToString();
            //        exCalc.ExportDataToEx(BaseTen[i], i * 11 + 1, 1, bt, fpath, 1);
            //    }

            //    //exCalc.wkten.Close(1);
            //}




            //获取对应赛道的10*10表

            //int[][,] BaseTen = exCalc.GetTenFive(exCalc.FilterData(dateTimePicker1.Value, dateTimePicker2.Value));

            //string fpath = string.Format(@"E:\制作\极速时时彩分析\测试文件目录\btExporttest{0:HHmmss}.xlsx", DateTime.Now);//设定导出文件名及路径;

            //exCalc.wkten = exCalc.CreateWk(fpath,4);
            //for (int i = 0; i < BaseTen.GetLength(0); i++)
            //{
            //    string bt = this.listBox1.Items[exCalc.LotSD[i]].ToString();
            //    exCalc.ExportDataToEx(BaseTen[i], i * 11 + 1, 1, bt,  1);
            //}



            //exCalc.wkten.Close(1);
            //wk1.Close(1);
            //string[] bt = new string[MaxNum];
            //bt[0] = "开奖号";
            //int[,] tempbt = new int[ MaxNum ,2];
            //for (int i = 1; i < MaxNum; i++)
            //{
            //    bt[i] =MaxNum>11 ? i.ToString():(i-1).ToString();
            //    tempbt[i - 1,0] = int.Parse(bt[i]);
            //}
            //exCalc.kjh = tempbt;
            //string fpath = string.Format(@"E:\制作\极速时时彩分析\测试文件目录\btExporttest{0:HHmmss}.xlsx", DateTime.Now);//设定导出文件名及路径
            //exCalc.ExportDataToEx(BaseTen[0], 1, 1, bt, fpath, 1);

            ////this.InitializeComponent();

            MessageBox.Show("Done!");
        }

        private void SingleDayCheck_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void SSC_CheckedChanged(object sender, EventArgs e)
        {
            LotCount = MultiDayCheck.Checked == true ? 10 : 5;
            AddIitemsForLBAndCB();
        }
    }
}
