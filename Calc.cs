using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization.Configuration;
using Excel = Microsoft.Office.Interop.Excel;

namespace 批量验证程序V1._0
{
    class Calc
    {
        /// <summary>
        /// 彩票开奖号码数，赛车10，时时彩5
        /// </summary>
        public int LotCount { get; set; }


        /// <summary>
        /// 开奖号码中的最大数。赛车为10，时时彩为9
        /// </summary>
        public int MaxNum { get; set; }

        /// <summary>
        /// 设定标题列名
        /// </summary>
        public string[] TitleColumn { get; set; }


        /// <summary>
        /// 基础数据表(10*10)
        /// </summary>
        public Workbook wkten { get; set; }


        /// <summary>
        /// 结果导出表
        /// </summary>
        public Workbook wkResult { get; set; }

        /// <summary>
        /// 获取标题行名及确定列名，赛车为1-10，时时彩0-9
        /// </summary>
        /// <param name="m">开奖号码中的最大数。赛车为10，时时彩为9</param>
        /// <returns></returns>
        public string[,] GetTitleLine(int m)
        {

            string[] Title5 = new string[] { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9" };

            string[] Title10 = new string[] { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10" };

            string[,] temparr = new string[10, 2];
            if (m>9)
            {
                for (int i = 1 ; i < 11; i++)
                {
                    temparr[i - 1, 0] = i.ToString();
                }
                TitleColumn = Title10;
            }
            else
            {
                for (int i = 0; i <10; i++)
                {
                    temparr[i, 0] = i.ToString();
                }
                TitleColumn = Title5;
            }
            return temparr;
        }

        /// <summary>
        /// 存放需要测试的赛道，从0开始的。
        /// </summary>
        public int[] LotSD;//存放需要测试的赛道



        /// <summary>
        /// 定义Excel工作薄相关的对象（workbook,worksheet,Excelapplication)
        /// </summary>
        public Excel.Application app = new Excel.Application();
   
        Workbook wk;
        Worksheet st;



        /// <summary>
        /// 过滤方式，Max就是过滤最大的，Min就是过滤最少的
        /// </summary>
        public string FilterType { get; set; }


        /// <summary>
        /// 过滤位数
        /// </summary>
        public int FilterCount { get; set; }


        /// <summary>
        /// Excel的最后一行，行号
        /// </summary>
        public int exEndRow { get; set; }
        

        /// <summary>
        /// 储存读取到的Excel数据
        /// </summary>
        public object[,] ExcelData { get; set; }


        /// <summary>
        /// 新建工作薄
        /// </summary>
        /// <param name="fpath">文件地址及文件名</param>
        /// <returns></returns>
        public Workbook CreateWk(string fpath,int stcount)
        {
            app.SheetsInNewWorkbook = stcount;
            Workbook wk1 = app.Workbooks.Add();
            //wk1.Sheets.Add();
            wk1.SaveAs(fpath);
            return wk1;
        }


        /// <summary>
        /// 打开工作薄
        /// </summary>
        /// <param name="fpath">文件名及地址</param>
        /// <returns></returns>
        public Workbook OpenWK(string fpath)
        {
            Workbook wk1 = app.Workbooks.Open(fpath);
            return wk1;
        }

        /// <summary>
        /// 获取并储存Excel地址
        /// </summary>
        public string ExFilePath{get;set;}
            
        public object[,] BaseData { get; set; }

        public object[,] testData { get; set; }

        public object[,] exFilterData { get; set; }


        public object[,] testData1 { get; set; }


        /// <summary>
        /// 读取Excel数据存入数组
        /// </summary>
        /// <param name="exFileName"></param>
        public void ReadExcel(string exFileName)
        {
            app.Visible = true;
            wk = app.Workbooks.Open(exFileName);
            st = wk.Sheets[1];
            exEndRow = st.UsedRange.Rows.Count;
            ExcelData = (object[,])st.UsedRange.Value;
            //BaseData=FilterData((DateTime)ExcelData[2,2],(DateTime)ExcelData[101,2]);
        }


        /// <summary>
        /// 筛选数据，通过com接口直接筛选Excel中的数据
        /// </summary>
        /// <param name="exFileName">Excel地址</param>
        /// <param name="startTime">开始日期及时间</param>
        /// <param name="endTime">结束日期及时间</param>
        /// <returns></returns>
        public object[,] FilterData(DateTime startTime, DateTime endTime)
        {
            st.UsedRange.AutoFilter(2, ">=" + startTime, XlAutoFilterOperator.xlAnd, "<=" + endTime, true);

            object[,] tempData= st.Range[st.Cells[2, 1], st.Cells[exEndRow, 3]].SpecialCells(XlCellType.xlCellTypeVisible).Value;

            st.UsedRange.AutoFilter(2);

            return tempData;
        }


        //public void test(DateTime stt,DateTime ent)
        //{
        //    object[,] tempData;
        //    st.UsedRange.AutoFilter(2, ">=" + stt, XlAutoFilterOperator.xlAnd, "<=" + ent, true);
        //    //testData = st.UsedRange.SpecialCells(XlCellType.xlCellTypeVisible).Value;

        //    tempData=st.Range[st.Cells[2,1],st.Cells[exEndRow,3]].SpecialCells(XlCellType.xlCellTypeVisible).Value;

        //    st.UsedRange.AutoFilter(2);

        //    return tempData;
        //    //ClearApp();

        //}


        //private int[,] GetTen()


        /// <summary>
        /// 生成10*10表
        /// </summary>
        /// <param name="arr"></param>
        public int[][,] GetTenFive(object[,] arr)
        {
            char[] spt = new char[] { ',' };
            int lcount = arr[1, 3].ToString().Split(spt).Length;/*这里与初始设定的号数理论上是一样为了避免人为错误因为直接从开奖号中读取位数。*/
            
            int[][,] TenFive = new int[LotSD.Length][,];
            for (int j = 0; j < LotSD.Length; j++)
            {

                int[,] temparr = new int[10, 10];
                for (int i = 1; i < arr.GetLength(0); i++)
                {
                    /*string[] temparr = arr[i, 3].ToString().Split(spt);
                    string[] temparr1 = arr[i + 1, 3].ToString().Split(spt);*/

                    int r = int.Parse(arr[i, 3].ToString().Split(spt)[LotSD[j]]);
                    //把10*10结果统计到TenFive数组中
                    int c = int.Parse(arr[i+1, 3].ToString().Split(spt)[LotSD[j]]);

                    temparr[r,c] += 1;

                }
                TenFive[j] = temparr;
            }
            return TenFive;
        }
        

        /// <summary>
        /// 根据10*10生成对应的筛选数据表。
        /// </summary>
        /// <param name="ten">对应赛道的10*10</param>
        /// <returns></returns>
        private int[,] GetFilterNum(int[,] ten)
        {
            int[,] FilterNums = new int[10, FilterCount];
            for (int i = 0; i < ten.GetLength(0); i++)
            {
                int[] temp = new int[10];
                int[] temparr = new int[10];
                

                for (int j = 0; j < ten.GetLength(1); j++)
                {
                    temparr[j] =ten[i, j];
                    temp[j] = j;
                }
                
                Array.Sort(temparr, temp);//从小到大的一个数组。

                //FilterType:Max ，Min
                if (FilterType=="Min")
                {
                    for (int j = 0; j <FilterCount; j++)
                    {
                        FilterNums[i, j] = temp[j];
                    }
                }
                else
                {
                    for (int j = temp.GetLength(0) - 1; j >= temp.GetLength(0)-FilterCount; j--)
                    {
                        FilterNums[i, temp.GetLength(0)-j-1] = temp[j];
                    }
                }


            }
            return FilterNums;
        }


        /// <summary>
        /// 用过滤数据测试目标数据，生成单独赛道的数据
        /// </summary>
        /// <param name="Dataarr"></param>
        /// <param name="Basearr"></param>
        /// <param name="TargetSD"></param>
        /// <returns></returns>
        private object[,] TestTargetWithBase(object[,] Dataarr,int[,] Basearr,int TargetSD)
        {
            int[] SDData = new int[Dataarr.GetLength(0)];
            object[,] temparr = new object[Dataarr.GetLength(0), 2];
            for (int i = 1; i < Dataarr.GetLength(0); i++)
            {
                SDData[i-1]= int.Parse(Dataarr[i, 3].ToString().Split(new char[] { ',' })[TargetSD]);
                SDData[i] = int.Parse(Dataarr[i+1, 3].ToString().Split(new char[] { ',' })[TargetSD]);
                //int r1 = SDData[i - 1];
                //int c1 = SDData[i];
                for (int j = 0; j < Basearr.GetLength(1); j++)
                {
                    if (Basearr[SDData[i-1],j]==SDData[i])
                    {
                        if (i==1)
                        {
                            temparr[i - 1, 0] = 1;
                        }
                        else
                        {
                            temparr[i - 1, 0] = (int)temparr[i - 2, 0] + 1;
                        }
                        break;
                    }
                    else
                    {
                        temparr[i - 1,0] =0;
                    }
                }

            }
            

            return temparr;
        }



        public object[,] CountTestResult(object[,] TargetArr)
        {

            object[,] CountArr = new object[500, 2];
            for (int i = 0; i < CountArr.GetLength(0); i++)
            {
                CountArr[i, 0] = 0;
            }
            for (int i = 0; i < TargetArr.GetLength(0)-1; i++)
            {
                
                CountArr[int.Parse(TargetArr[i, 0].ToString()), 0] = (int)CountArr[int.Parse(TargetArr[i, 0].ToString()), 0]+ 1;
            }
            return CountArr;
        }


        /// <summary>
        /// 导出数组到Excel中。本方法操作基础数据汇总及筛选结果导出。
        /// </summary>
        /// <param name="arr">需要导出的数组</param>
        /// <param name="Frow">导出的位置所在行</param>
        /// <param name="Fcol">导出位置所在列</param>
        /// <param name="Flines">标题行内容</param>
        /// <param name="fpath">地址及文件名</param>
        /// <param name="SheetIndex">导出的工作表的索引</param>
        public void ExportDataToEx(int[,] arr1,int Frow,int Fcol,string bt,int SheetIndex)
        {
            //Workbook wks = File.Exists(fpath) ? OpenWK(fpath) : CreateWk(fpath);

            Worksheet wst = wkten.Sheets[SheetIndex];
            string[,] TitleRow =GetTitleLine(MaxNum);
            int ColCount = arr1.GetLength(1);
            wst.Cells[Frow, Fcol].Value =bt;
            wst.Range[wst.Cells[Frow, Fcol +1 ], wst.Cells[Frow,Fcol+ColCount]].Value =TitleColumn;//TitleColumn
            wst.Range[wst.Cells[Frow+1, Fcol], wst.Cells[Frow+10, Fcol]].Value = TitleRow;//TitleRow
            wst.Range[wst.Cells[Frow+1,Fcol+1], wst.Cells[Frow+10, Fcol+ColCount]].Value = arr1;
            wst.Cells.EntireColumn.AutoFit();
            //wst.Cells.HorizontalAlignment = 2;
            //wst.Cells.VerticalAlignment = 2;
            //wst.Cells.EntireColumn.HorizontalAlignment=HorizontalAlignment.Center;
            //wst.Name = SheetIndex.ToString();
            //wks.Close(1);

            //wks.Close(1);

            //ClearApp();
        }



        /// <summary>
        /// 导出数组到Excel中。本方法操作导出测试结果及统计结果。
        /// </summary>
        /// <param name="arr">需要导出的数组</param>
        /// <param name="Frow">导出的位置所在行</param>
        /// <param name="Fcol">导出位置所在列</param>
        /// <param name="Flines">标题行内容</param>
        /// <param name="fpath">地址及文件名</param>
        /// <param name="SheetIndex">导出的工作表的索引</param>
        public void ExportDataToEx(object[,] arr, int Frow, int Fcol, string bt, int SheetIndex)
        {
            Worksheet wst = wkResult.Sheets[SheetIndex];
            if (Fcol==1)
            {
                string[,] TitleRow = GetTitleLine(MaxNum);
                int ColCount = arr.GetLength(1);
                int RowCount = arr.GetLength(0);
                wst.Cells[Frow, Fcol].Value = bt;

                //wst.Range[wst.Cells[Frow, Fcol + 1], wst.Cells[Frow, Fcol + ColCount]].Value = TitleColumn;//TitleColumn
                //wst.Range[wst.Cells[Frow + 1, Fcol], wst.Cells[Frow + 10, Fcol]].Value = TitleRow;//TitleRow
                wst.Range[wst.Cells[Frow + 1, 1], wst.Cells[Frow + RowCount,  ColCount]].Value = arr;
                wst.Cells.EntireColumn.AutoFit();
                wst.Range[wst.Cells[2, 2], wst.Cells[RowCount+1, 2]].NumberFormatLocal = "yyyy/mm/dd hh:mm";
            }
            else
            {
                wst.Cells[1, Fcol] = bt;
                wst.Range[wst.Cells[2, Fcol], wst.Cells[1+arr.GetLength(0),Fcol]].Value = arr;

            }



        }



        public void GetBaseFilterData(bool sm,DateTime t1,DateTime t2,DateTime t3,DateTime t4,int filterNum,int cycleNum)
        {
            //基础数据单日模式是周期内一天一个结果，多日模式是周期内只有一个结果。

            //Excel.Worksheet stt = (Excel.Worksheet)wkten.Sheets.get_Item(wkten.Sheets.Count);
            //Excel._Worksheet st1 = (Excel._Worksheet)wkten.Sheets.Add(Missing.Value, stt, 4, XlSheetType.xlWorksheet);

            //wkten = app.Workbooks.Add();

            //wkten.SaveAs(fpath);

            //Workbook wkk = app.Workbooks[2];
            //MessageBox.Show(wkk.Name);
            string basePath = string.Format(@"E:\制作\极速时时彩分析\测试文件目录\BaseData{0:HHmmss}.xlsx",
                DateTime.Now);
            string TestPath = string.Format(@"E:\制作\极速时时彩分析\测试文件目录\ResultData{0:HHmmss}.xlsx",
                DateTime.Now);
            if (sm)
            {
                wkten = CreateWk(basePath, cycleNum);//新建基础数据表

                wkResult = CreateWk(TestPath, cycleNum);//新建测试数据表
                for (int i = 0; i < cycleNum; i++)
                {
                    object[,] TargetData = FilterData(t3, t4);


                    int[][,] BaseTen = GetTenFive(FilterData(t1, t2));//单日的基础数据是移到的


                    //目标文件保存到Exceldata中。用FilterData提取测试数据
                    ExportDataToEx(TargetData, 1, 1, "", i + 1);

                    //按赛道循环,基础数据，测试数据都是按赛道单独计算。最终统一生成。
                    for (int j = 0; j < BaseTen.GetLength(0); j++)
                    {
                        int[,] FilterArr = GetFilterNum(BaseTen[j]);
                        //导出赛道统计数据
                        ExportDataToEx(BaseTen[j], j * 11 + 1, 1, string.Format("第{0:D2}赛道10*10", LotSD[j] + 1), i + 1);

                        //导出过滤数统计结果
                        ExportDataToEx(FilterArr, j * 11 + 1, 13, string.Format("第{0:D2}赛道过滤数", LotSD[j] + 1), i + 1);

                        ExportDataToEx(TestTargetWithBase(TargetData, FilterArr, LotSD[j]), 2, j + 4, string.Format("第{0}赛道", LotSD[j] + 1), i + 1);


                    }

                    wkten.Worksheets[i + 1].Name = (i + 1).ToString();//修改sheet名称为数字编号基础数据表10*10
                    wkResult.Worksheets[i + 1].Name = (i + 1).ToString();//修改sheet名称为数字编号 测试数据表
                    Worksheet wws = wkResult.Worksheets[i + 1];
                    wws.Range[wws.Cells[1, 1], wws.Cells[1, 3]] = new string[] { "期号", "时间", "开奖号" };

                    t1 = t1.AddDays(1);
                    t2 = t2.AddDays(1);
                    t3 = t3.AddDays(1);
                    t4 = t4.AddDays(1);
                }
            }
            else
            {
                wkten = CreateWk(basePath, 1);//新建基础数据表
                wkResult = CreateWk(TestPath, cycleNum);//新建测试数据表

                int[][,] BaseTen = GetTenFive(FilterData(t1, t2));//多日的基础数据是固定的
                                                                  //因此不用循环
                                                                  //string fpath = string.Format(@"E:\制作\极速时时彩分析\测试文件目录\btExporttest{0:HHmmss}.xlsx", DateTime.Now);//设定导出文件名及路径;


                int[][,] FilterArr = new int[LotSD.GetLength(0)][,];
                //按赛道循环,基础数据，测试数据都是按赛道单独计算。最终统一生成。
                for (int i = 0; i < BaseTen.GetLength(0); i++)
                {
                    //string bt = this.listBox1.Items[exCalc.LotSD[i]].ToString();

                    //导出赛道统计数据
                    ExportDataToEx(BaseTen[i], i * 11 + 1, 1, string.Format("第{0:D2}赛道10*10", LotSD[i] + 1), 1);

                    //导出过滤数统计结果
                    ExportDataToEx(GetFilterNum(BaseTen[i]), i * 11 + 1, 13, string.Format("第{0:D2}赛道过滤数", LotSD[i] + 1), 1);
                    FilterArr[i] = GetFilterNum(BaseTen[i]);
                    //int[,] FilterArr = GetFilterNum(BaseTen[i]);
                    //object[,] TargetData = FilterData(t3, t4);

                    //for (int j = 0; j < cycleNum; j++)
                    //{

                    //    ExportDataToEx(TargetData, 1, 1, "", i + 1);
                    //    ExportDataToEx(TestTargetWithBase(TargetData, FilterArr, LotSD[i]), 2, j + 4, string.Format("第{0}赛道", LotSD[j] + 1), i + 1);

                    //}
                    //Worksheet wws = wkResult.Worksheets[i + 1];
                    //wws.Range[wws.Cells[1, 1], wws.Cells[1, 3]] = new string[] { "期号", "时间", "开奖号" };
                }
                string[,] RowTitle = new string[500, 2];
                RowTitle[0, 0] = "挂次数";
                for (int i = 1; i < RowTitle.GetLength(0); i++)
                {
                    RowTitle[i, 0] = i.ToString();
                }

                for (int i = 0; i < cycleNum; i++)
                {
                    Worksheet wws = wkResult.Worksheets[i + 1];
                    object[,] TargetData = FilterData(t3, t4);
                    ExportDataToEx(TargetData, 1, 1, "", i + 1);
                    wws.Range[wws.Cells[1, 14], wws.Cells[500, 14]] = RowTitle;

                    for (int j = 0; j < BaseTen.GetLength(0); j++)
                    {
                        object[,] ResultArr = TestTargetWithBase(TargetData, FilterArr[j], LotSD[j]);
                        ExportDataToEx(ResultArr,2,j+4, string.Format("第{0}赛道", LotSD[j] + 1), i + 1);
                        ExportDataToEx(CountTestResult(ResultArr),2,j+15, string.Format("第{0}赛道", LotSD[j] + 1), i + 1);
                    }


                    wws.Name = (i + 1).ToString();//修改sheet名称为数字编号 测试数据表
                    wws.Range[wws.Cells[1, 1], wws.Cells[1, 3]] = new string[] { "期号", "时间", "开奖号" };
                    t3 = t3.AddDays(1);
                    t4 = t4.AddDays(1);
                }

                //多日测试只有一个基础数据因为不用添加多张sheet。


            }
            //这里已经得到基础数据了。工作薄不用关闭，后面接着存入过滤数据。


        }


        /// <summary>
        /// 引用完成后清空各项引用
        /// </summary>
        public void ClearApp()
        {
            wk.Close(false);
            app.Quit();
            releaseObject(st);
            releaseObject(wk);
            releaseObject(app);
        }


        /// <summary>
        /// 清空引用，清理后台应用，垃圾回收。
        /// </summary>
        /// <param name="obj"></param>
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (System.Exception ex)
            {
                obj = null;
                //MessageBox.Show("Unable to release the Object" + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }

        }


    }
}
