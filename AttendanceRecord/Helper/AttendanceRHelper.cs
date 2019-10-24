using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Tools;
using AttendanceRecord.Entities;
using Excel;
using Oracle.DataAccess.Client;
using System.Data;
namespace AttendanceRecord.Helper
{
    public static class AttendanceRHelper
    {
        static string tempStr = String.Empty;
        /// <summary>
        /// 将考勤记录导入数据库.
        /// </summary>
        /// <param name="xlsFilePath"></param>
        /// <param name="randomStr"></param>
        /// <param name="pb"></param>
        /// <returns></returns>
        public static MSG  ImportAttendanceRecordToDB(string xlsFilePath, string randomStr, System.Windows.Forms.Label lblPrompt,ProgressBar pb,System.Windows.Forms.Label lblResult)
        {
            int pbLength = 0;
            MSG msg = new MSG();

            //用于确定本月最后一天.
            Stack<int> sDate = new Stack<int>();
            Queue<AttendanceR> qAttendanceR = new Queue<AttendanceR>();
            AttendanceR.Random_Str = randomStr;
            //按指纹日期
            string fingerPrintDate = String.Empty ;
            //导入数据的行数.
            int affectedCount = 0;
            //行最大值.
            int rowsMaxCount = 0;
            int colsMaxCount = 0;
            Usual_Excel_Helper uEHelper = null;
            MyExcel myExcel = new MyExcel(xlsFilePath);
            //打开该文档。
            myExcel.open();
            #region 获取该工作簿中，行数最少的表格。
            #endregion
            //只获取第一个表格。
            Worksheet ws = myExcel.getFirstWorkSheetAfterOpen();
            AttendanceR.File_path = xlsFilePath;
            //行;列最大值 赋值.
            rowsMaxCount = ws.UsedRange.Rows.Count;
            colsMaxCount = ws.UsedRange.Columns.Count;
            pb.Value = 0;
            pb.Visible = true;
            lblPrompt.Visible = true;
            AttendanceR.Sheet_name = ws.Name;
            //判断首行是否为 考勤记录表;以此判断此表是否为考勤记录表.
            string A1Str =((Range)ws.Cells[1, 1]).Text.ToString().Trim().Replace("\n","").Replace("\r","").Replace(" ","");
            if (String.IsNullOrEmpty(A1Str)) {
                msg.Msg = "工作表的A1单元格不能为空！";
                msg.Flag = false;
                myExcel.close();
                return msg;
            }
            //如果A1Str的内容不包含"考勤记录表"5个字。       
            if (!A1Str.Contains("考勤记录表")) {
                msg.Msg = "A1内容未包含'考勤记录表'";
                msg.Flag = false;
                myExcel.close();
                return msg;        
            }
            #region 判断名称中是否区分了考勤记录。
            string  Seq_Attendance_Record = string.Empty;
            int indexOfFullStop = xlsFilePath.LastIndexOf(".");
            Seq_Attendance_Record = xlsFilePath.Substring(indexOfFullStop - 1, 1);
            if (!CheckPattern.CheckNumber(Seq_Attendance_Record))
            {
                msg.Msg = "考勤记录表名称请以数字结尾！";
                msg.Flag = false;
                myExcel.close();
                return msg;
            }
            #endregion
            string excelName = Usual_Excel_Helper.getExcelName(xlsFilePath);
            AttendanceR.Prefix_Job_Number = excelName.Substring(excelName.Length - 1, 1).ToCharArray()[0];
            string C3Str = ((Range)ws.Cells[3, 3]).Text.ToString().Trim();
            //  \0: 表空字符.
            if (String.IsNullOrEmpty(C3Str)) {
                msg.Msg = "异常: 考勤时间为空!";
                msg.Flag = false;
                myExcel.close();
                return msg;
            }
            //
            string[] ArrayC3 = C3Str.Split('~');
            if (ArrayC3.Length == 0) {
                msg.Msg = "异常: 考勤时间格式变更!";
                msg.Flag = false;
                myExcel.close();
                return msg;
            }
            AttendanceR.Start_date= ArrayC3[0].ToString().Trim();
            AttendanceR.End_date = ArrayC3[1].ToString().Trim();
            //制表时间:  L3 3行12列.
            string L3Str =((Range) ws.Cells[3, 12]).Text.ToString().Trim();
            if (String.IsNullOrEmpty(L3Str))
            {
                msg.Msg = "异常: 制表时间为空!";
                msg.Flag = false;
                myExcel.close();
                return msg;
            }
            //制表时间.
            AttendanceR.Tabulation_time = L3Str;
            //检查第4行是否为;考勤时间:
            string A4Str =((Range) ws.Cells[4, 1]).Text.ToString().Trim();
            if (!"1".Equals(A4Str, StringComparison.CurrentCultureIgnoreCase)) {
                msg.Msg = "异常: 第四行已变更!";
                msg.Flag = false;
                myExcel.close();
                return msg;
            }
            uEHelper = new Usual_Excel_Helper(ws);
            //此刻不能删除，只是获取行号。
            Queue<Range> rangeToDelQueue = new Queue<Range>();
            //判断是否有空行。
            for (int i = 5;i<=rowsMaxCount;i++) {
                if (uEHelper.isBlankRow(i)) {
                    //只要上一列不是
                    //删除掉此行。
                    //判断上一行中的A列是否为工号。
                    string temp = uEHelper.getSpecificCellValue("A" + (i - 1).ToString());
                    if ("工号:".Equals(temp)) {
                        //本行为空，上一行为工号行，则也统计。
                        continue;
                    }
                    //本行，为空，上一行非工号行。则删除本行。
                    Range rangeToDel = (Microsoft.Office.Interop.Excel.Range)uEHelper.WS.Rows[i, System.Type.Missing];
                    //不为工号
                    rangeToDelQueue.Enqueue(rangeToDel);
                };
            }
            Range rangeToDelete;
            //开始删除空行。  
            while (rangeToDelQueue.Count>0) {
                rangeToDelete = rangeToDelQueue.Dequeue();
                rangeToDelete.Delete(XlDeleteShiftDirection.xlShiftUp);
            } ;
            rowsMaxCount = ws.UsedRange.Rows.Count;
            //进度条长度增加。
            pbLength += colsMaxCount;
            pbLength += (colsMaxCount * (rowsMaxCount - 5 + 1));
            pb.Maximum = pbLength;
            //入队列值0
            sDate.Push(0);
            //显示进度条。
            //考勤表中第4行，某月的最大考勤天数。
            lblPrompt.Text = excelName + "，正在读取：";
            int actualMaxDay = 0;
            //开始循环
            for (int i=1;i<=colsMaxCount;i++) {
                A4Str = ((Range)ws.Cells[4, i]).Text.ToString();
                //碰到第4行某列为空，退出循环。
                if (String.IsNullOrEmpty(A4Str))
                {
                    break;
                }
                int aDate = 0;
                //对A4Str进行分析.
                if (!Int32.TryParse(A4Str, out aDate))
                {
                    msg.Msg =String.Format(@"异常: 考勤日期行第{0}列出现非数字内容!", aDate);
                    msg.Flag = false;
                    myExcel.close();
                    return msg;
                }
                pb.Value++;
                //判断新增的日期是否大于上一个.
                if (aDate <= sDate.Peek()) {
                    //跳出循环.
                    break;
                }
                actualMaxDay++;
                sDate.Push(aDate);
            }
            //取其中的最小值。
            colsMaxCount = Math.Min(sDate.Count-1,actualMaxDay);
            //考勤日期
            fingerPrintDate = AttendanceR.Start_date.Substring(0,7).Replace('/','-');
            string tempStr = string.Empty;
            //开始循环
            for (int colIndex= 1;colIndex <=colsMaxCount;colIndex ++) {
                //从第5行开始.
                //奇数;偶数行共用一个对象.
                AttendanceR AR = null;  
                for (int rowIndex = 5;rowIndex <= rowsMaxCount;rowIndex ++) {
                    //如果行数为奇数则为工号行.
                    if (rowIndex % 2 == 1)
                    {
                        //工号行.
                        //取工号
                        AR = new AttendanceR();
                        AR.Job_number = ((Range)ws.Cells[rowIndex, 3]).Text.ToString().Trim();
                        //自行拼凑AR.
                        AR.combine_Job_Number();
                        //取姓名:  K5 
                        AR.Name =((Range)ws.Cells[rowIndex, Usual_Excel_Helper.getColIndexByStr("K")]).Text.ToString().Trim();
                        //取部门: U5
                        AR.Dept = ((Range) ws.Cells[rowIndex, Usual_Excel_Helper.getColIndexByStr("U")]).Text.ToString().Trim();
                        //部门为空，则填充为NULL;
                        AR.Dept = !String.IsNullOrEmpty(AR.Dept) ? AR.Dept : "NULL";
                        //取日期.填充0;
                        AR.Fingerprint_Date = fingerPrintDate + "-" + colIndex.ToString().PadLeft(2, '0');
                    }
                    else {
                        //偶数行取考勤结果.
                        //上班时间. 如B10;
                        tempStr = ((Range)ws.Cells[rowIndex, colIndex]).Text.ToString().Trim();
                        string tempFirstTime = String.Empty;
                        string tempLastTime = String.Empty;
                        if (!getFPTime(tempStr, out tempFirstTime, out tempLastTime)) {
                            msg.Msg = string.Format(@"导入失败：表中第{0}行{1}列的按指纹时间格式不对！", rowIndex, colIndex);
                            msg.Flag = false;
                            myExcel.close();
                            return msg;
                        } ;
                        AR.FPT_Fisrt_Time = String.IsNullOrEmpty(tempFirstTime) ? String.Empty : AR.Fingerprint_Date + " " + tempFirstTime;
                        AR.FPT_Last_Time = String.IsNullOrEmpty(tempLastTime)? String.Empty : AR.Fingerprint_Date + " " + tempLastTime;
                        qAttendanceR.Enqueue(AR);
                    }
                    pb.Value++;
                }
            }
            //释放对象
            myExcel.close();
            System.Threading.Thread.Sleep(2000);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            lblResult.Text = "";
            lblResult.Visible = false;
            lblPrompt.Visible = false;
            lblPrompt.Text = "";
            lblPrompt.Text = "提交数据: ";
            lblPrompt.Focus();
            lblPrompt.Visible = true;
            pb.Maximum = qAttendanceR.Count;
            pb.Value = 0;
            #region
            OracleDaoHelper.noLogging("Attendance_Record");
            OracleConnection conn = OracleConnHelper.getConn();
            OracleTransaction tran = conn.BeginTransaction();
            //保存对象
            while (qAttendanceR.Count > 0) {
                try
                {
                    AttendanceR aR= qAttendanceR.Dequeue();
                    affectedCount += aR.importAR(conn);
                    pb.Value++;
                }
                catch (Exception ex) {
                    MessageBox.Show(ex.ToString(), "提示:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    msg.Msg = DirectoryHelper.getFileName(xlsFilePath) + "：导入失败; " + ex.ToString() ;
                    msg.Flag = false;
                    tran.Rollback();
                    conn.Close();
                    conn.Dispose();
                    return msg;
                }
            }
            tran.Commit();
            conn.Close();
            conn.Dispose();
            #endregion
            OracleDaoHelper.logging("Attendance_Record");
            lblPrompt.Visible = false;
            //隐藏进度条。
            pb.Visible = false;
            msg.Flag = true;
            msg.Msg = String.Format(@"导入完成;计{0}条.", affectedCount.ToString());
            return msg;
        }
        /// <summary>
        /// 将考勤记录导入预备表中.
        /// </summary>
        /// <param name="xlsFilePath"></param>
        /// <param name="randomStr"></param>
        /// <param name="pb"></param>
        /// <returns></returns>
        public static MSG Import_Record_To_Preparative_Table(string xlsFilePath, string randomStr, System.Windows.Forms.Label lblPrompt, ProgressBar pb, System.Windows.Forms.Label lblResult)
        {
            int pbLength = 0;
            MSG msg = new MSG();

            //用于确定本月最后一天.
            Stack<int> sDate = new Stack<int>();
            Queue<AttendanceR> qAttendanceR = new Queue<AttendanceR>();
            AttendanceR.Random_Str = randomStr;
            //按指纹日期
            string fingerPrintDate = String.Empty;
            //导入数据的行数.
            int affectedCount = 0;
            //行最大值.
            int rowsMaxCount = 0;
            int colsMaxCount = 0;
            Usual_Excel_Helper uEHelper = null;
            MyExcel myExcel = new MyExcel(xlsFilePath);
            //打开该文档。
            myExcel.open();
         
            //只获取第一个表格。
            Worksheet ws = myExcel.getFirstWorkSheetAfterOpen();
            AttendanceR.File_path = xlsFilePath;
            //行;列最大值 赋值.
            rowsMaxCount = ws.UsedRange.Rows.Count;
            colsMaxCount = ws.UsedRange.Columns.Count;
            pb.Value = 0;
            pb.Visible = true;
            lblPrompt.Visible = true;
            AttendanceR.Sheet_name = ws.Name;
            //判断首行是否为 考勤记录表;以此判断此表是否为考勤记录表.
            string A1Str = ((Range)ws.Cells[1, 1]).Text.ToString().Trim().Replace("\n", "").Replace("\r", "").Replace(" ", "");
            if (String.IsNullOrEmpty(A1Str))
            {
                msg.Msg = "工作表的A1单元格不能为空！";
                msg.Flag = false;
                myExcel.close();
                return msg;
            }
            //如果A1Str的内容不包含"考勤记录表"5个字。       
            if (!A1Str.Contains("考勤记录表"))
            {
                msg.Msg = "A1内容未包含'考勤记录表'";
                msg.Flag = false;
                myExcel.close();
                return msg;
            }
            #region 判断名称中是否区分了考勤记录。
            string Seq_Attendance_Record = string.Empty;
            int indexOfFullStop = xlsFilePath.LastIndexOf(".");
            Seq_Attendance_Record = xlsFilePath.Substring(indexOfFullStop - 1, 1);
            if (!CheckPattern.CheckNumber(Seq_Attendance_Record))
            {
                msg.Msg = "考勤记录表名称请以数字结尾！";
                msg.Flag = false;
                myExcel.close();
                return msg;
            }
            #endregion
            string excelName = Usual_Excel_Helper.getExcelName(xlsFilePath);
            AttendanceR.Prefix_Job_Number = excelName.Substring(excelName.Length - 1, 1).ToCharArray()[0];
            string C3Str = ((Range)ws.Cells[3, 3]).Text.ToString().Trim();
            //  \0: 表空字符.
            if (String.IsNullOrEmpty(C3Str))
            {
                msg.Msg = "异常: 考勤时间为空!";
                msg.Flag = false;
                myExcel.close();
                return msg;
            }
            //
            string[] ArrayC3 = C3Str.Split('~');
            if (ArrayC3.Length == 0)
            {
                msg.Msg = "异常: 考勤时间格式变更!";
                msg.Flag = false;
                myExcel.close();
                return msg;
            }
            AttendanceR.Start_date = ArrayC3[0].ToString().Trim();
            AttendanceR.End_date = ArrayC3[1].ToString().Trim();
            //制表时间:  L3 3行12列.
            string L3Str = ((Range)ws.Cells[3, 12]).Text.ToString().Trim();
            if (String.IsNullOrEmpty(L3Str))
            {
                msg.Msg = "异常: 制表时间为空!";
                msg.Flag = false;
                myExcel.close();
                return msg;
            }
            //制表时间.
            AttendanceR.Tabulation_time = L3Str;
            //检查第4行是否为;考勤时间:
            string A4Str = ((Range)ws.Cells[4, 1]).Text.ToString().Trim();
            if (!"1".Equals(A4Str, StringComparison.CurrentCultureIgnoreCase))
            {
                msg.Msg = "异常: 第四行已变更!";
                msg.Flag = false;
                myExcel.close();
                return msg;
            }
            uEHelper = new Usual_Excel_Helper(ws);
            //此刻不能删除，只是获取行号。
            Queue<Range> rangeToDelQueue = new Queue<Range>();
            //判断是否有空行。
            for (int i = 5; i <= rowsMaxCount; i++)
            {
                if (uEHelper.isBlankRow(i))
                {
                    //只要上一列不是
                    //删除掉此行。
                    //判断上一行中的A列是否为工号。
                    string temp = uEHelper.getSpecificCellValue("A" + (i - 1).ToString());
                    if ("工号:".Equals(temp))
                    {
                        //本行为空，上一行为工号行，则也统计。
                        continue;
                    }
                    //本行，为空，上一行非工号行。则删除本行。
                    Range rangeToDel = (Microsoft.Office.Interop.Excel.Range)uEHelper.WS.Rows[i, System.Type.Missing];
                    //不为工号
                    rangeToDelQueue.Enqueue(rangeToDel);
                };
            }
            Range rangeToDelete;
            //开始删除空行。  
            while (rangeToDelQueue.Count > 0)
            {
                rangeToDelete = rangeToDelQueue.Dequeue();
                rangeToDelete.Delete(XlDeleteShiftDirection.xlShiftUp);
            };
            rowsMaxCount = ws.UsedRange.Rows.Count;
            //进度条长度增加。
            pbLength += colsMaxCount;
            pbLength += (colsMaxCount * (rowsMaxCount - 5 + 1));
            pb.Maximum = pbLength;
            //入队列值0
            sDate.Push(0);
            //显示进度条。
            //考勤表中第4行，某月的最大考勤天数。
            lblPrompt.Text = excelName + "，正在读取：";
            int actualMaxDay = 0;
            //开始循环
            for (int i = 1; i <= colsMaxCount; i++)
            {
                A4Str = ((Range)ws.Cells[4, i]).Text.ToString();
                //碰到第4行某列为空，退出循环。
                if (String.IsNullOrEmpty(A4Str))
                {
                    break;
                }
                int aDate = 0;
                //对A4Str进行分析.
                if (!Int32.TryParse(A4Str, out aDate))
                {
                    msg.Msg = String.Format(@"异常: 考勤日期行第{0}列出现非数字内容!", aDate);
                    msg.Flag = false;
                    myExcel.close();
                    return msg;
                }
                pb.Value++;
                //判断新增的日期是否大于上一个.
                if (aDate <= sDate.Peek())
                {
                    //跳出循环.
                    break;
                }
                actualMaxDay++;
                sDate.Push(aDate);
            }
            //取其中的最小值。
            colsMaxCount = Math.Min(sDate.Count - 1, actualMaxDay);
            //考勤日期
            fingerPrintDate = AttendanceR.Start_date.Substring(0, 7).Replace('/', '-');
            string tempStr = string.Empty;
            //开始循环
            for (int colIndex = 1; colIndex <= colsMaxCount; colIndex++)
            {
                //从第5行开始.
                //奇数;偶数行共用一个对象.
                AttendanceR AR = null;
                for (int rowIndex = 5; rowIndex <= rowsMaxCount; rowIndex++)
                {
                    //如果行数为奇数则为工号行.
                    if (rowIndex % 2 == 1)
                    {
                        //工号行.
                        //取工号
                        AR = new AttendanceR();
                        AR.Job_number = ((Range)ws.Cells[rowIndex, 3]).Text.ToString().Trim();
                        //自行拼凑AR.
                        AR.combine_Job_Number();
                        //取姓名:  K5 
                        AR.Name = ((Range)ws.Cells[rowIndex, Usual_Excel_Helper.getColIndexByStr("K")]).Text.ToString().Trim();
                        //取部门: U5
                        AR.Dept = ((Range)ws.Cells[rowIndex, Usual_Excel_Helper.getColIndexByStr("U")]).Text.ToString().Trim();
                        //部门为空，则填充为NULL;
                        AR.Dept = !String.IsNullOrEmpty(AR.Dept) ? AR.Dept : "NULL";
                        //取日期.填充0;
                        AR.Fingerprint_Date = fingerPrintDate + "-" + colIndex.ToString().PadLeft(2, '0');
                    }
                    else
                    {
                        //偶数行取考勤结果.
                        //上班时间. 如B10;
                        tempStr = ((Range)ws.Cells[rowIndex, colIndex]).Text.ToString().Trim();
                        string tempFirstTime = String.Empty;
                        string tempLastTime = String.Empty;
                        if (!getFPTime(tempStr, out tempFirstTime, out tempLastTime))
                        {
                            msg.Msg = string.Format(@"导入失败：表中第{0}行{1}列的按指纹时间格式不对！", rowIndex, colIndex);
                            msg.Flag = false;
                            myExcel.close();
                            return msg;
                        };
                        AR.FPT_Fisrt_Time = String.IsNullOrEmpty(tempFirstTime) ? String.Empty : AR.Fingerprint_Date + " " + tempFirstTime;
                        AR.FPT_Last_Time = String.IsNullOrEmpty(tempLastTime) ? String.Empty : AR.Fingerprint_Date + " " + tempLastTime;
                        qAttendanceR.Enqueue(AR);
                    }
                    pb.Value++;
                }
            }
            //释放对象
            myExcel.close();
            System.Threading.Thread.Sleep(2000);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            lblResult.Text = "";
            lblResult.Visible = false;
            lblPrompt.Visible = false;
            lblPrompt.Text = "";
            lblPrompt.Text = "提交数据: ";
            lblPrompt.Focus();
            lblPrompt.Visible = true;
            pb.Maximum = qAttendanceR.Count;
            pb.Value = 0;
            #region
            //保存对象
            while (qAttendanceR.Count > 0)
            {
                try
                {
                    AttendanceR aR = qAttendanceR.Dequeue();
                    affectedCount += aR.import_AR_To_Preparative_Table();
                    pb.Value++;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), "提示:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    msg.Msg = DirectoryHelper.getFileName(xlsFilePath) + "：导入失败; " + ex.ToString();
                    msg.Flag = false;
                    return msg;
                }
            }
            #endregion
            lblPrompt.Visible = false;
            //隐藏进度条。
            pb.Visible = false;
            msg.Flag = true;
            msg.Msg = String.Format(@"导入完成;计{0}条.", affectedCount.ToString());
            return msg;
        }
        /// <summary>
        /// 目前只考虑白班.
        /// </summary>
        /// <param name="strTime"></param>
        /// <param name="FirstTime"></param>
        /// <param name="lastTime"></param>
        public static bool getFPTime(string strTime, out string firstTime, out string lastTime) {
            int hour = 0;
            int minute = 0;
            //目的是再次到处汇总的表格，因为汇总的表格中时间纪录以 \r\n 换行相隔。
            strTime = strTime.Replace("\r\n", "");
            strTime = strTime.Replace(" ","");
            strTime = strTime.Replace("  ", "");
            strTime = strTime.Replace("   ", "");
            //判断长度
            if (String.IsNullOrEmpty(strTime))
            {
                firstTime = "";
                lastTime = "";
                return true;
            }
            if (strTime.Substring(1, 1) == ":") {
                //如7:07 --> 07:07
                strTime = "0" + strTime;
            }
            //判断长度是否可以被5整除
            if (strTime.Length % 5 != 0) {
                firstTime = "";
                lastTime = "";
                return false;
            }
            List<string> timeStrList = new List<string>();
            for (int i=0;i<= strTime.Length/5 -1;i++) {
                timeStrList.Add(strTime.Substring(i*5, 5));
            }
            //排序好的字符串。
            List<DateTime> dtSortedList = new List<DateTime>();
            for (int i=0;i<= timeStrList.Count-1;i++) {
                //如果 时间字符串格式不符合规定，return false;
                bool flag = false;
                DateTime dt ;
                flag = DateTime.TryParse(timeStrList[i],out dt);
                if (!flag) {
                    firstTime = "";
                    lastTime = "";
                    return false;
                }
                dtSortedList.Add(dt);
            }
            //排序，再次输出。
            dtSortedList.Sort();
            strTime = string.Empty;
            for (int i = 0; i <= dtSortedList.Count - 1; i++) {
                strTime += dtSortedList[i].ToString("HH:mm");

            }
            #region 如果时间字符长度为5; 表示只刷了一次卡.
            if (strTime.Length == 5) {
                //判断是否为:  6 --> 11点
                hour = 0;
                minute = 0;
                //判断是否>=0:01 And< 12:20  
                int.TryParse(strTime.Substring(0, 2), out hour);
                int.TryParse(strTime.Substring(3, 2), out minute);
                if (hour >= 6 && (hour < 11))
                {
                    firstTime = strTime;
                    lastTime = "";
                    return true;
                }
                else if (hour<6 && hour >=0) {
                    //凌晨间的刷卡，也 计为：早上卡。
                    firstTime = strTime;
                    lastTime = "";
                    return true;
                }
                else {
                    firstTime = "";
                    lastTime = strTime;
                    return true;
                }
            }
            #endregion
            #region 刷卡次数为2次，判断第一次和第二次的间隔时间是否小于10分钟。
            //若小于10分钟，则认为Fpt_last_time, 不填写。
            if (10 == strTime.Length) {
                //获取
                string str1 = strTime.Substring(0, 5);
                string str2 = strTime.Substring(5, 5);

                DateTime dt1 = DateTime.Parse(str1);
                DateTime dt2 = DateTime.Parse(str2);
                double differentValue = (dt2 - dt1).TotalMinutes;
                if (differentValue < 10) {
                    firstTime = str1;
                    lastTime = "";
                    return true;
                }
                firstTime = str1;
                lastTime = str2;
                return true;
            }
            #endregion
            #region 刷卡次数:  2次以上.
            firstTime = strTime.Substring(0, 5);
            lastTime = strTime.Substring(strTime.Length - 5, 5);
            #endregion
            return true;
        }
        #region 检查时间内容是否正确
        public static MSG checkTimeStr(string timStr) {
            MSG msg = new MSG();
            //判断是否为5的倍数.
            if (timStr.Length % 5 != 0) {
                msg.Flag = false;
                msg.Msg = "时间长度不为5的倍数!";
                return msg;
            }
            return msg;
        }
        #endregion
        /// <summary>
        /// 将考勤记录导入数据库.
        /// </summary>
        /// <param name="xlsFilePath"></param>
        /// <param name="randomStr"></param>
        /// <param name="pb"></param>
        /// <returns></returns>
        public static MSG  ImportARToDB(string xlsFilePath)
        {
            int pbLength = 0;
            MSG msg = new MSG();
            ProgressInfo pI = new ProgressInfo("ConsoleCollectAR");
            if (pI.ifExists()) pI.delete();
            pI.add();
            //用于确定本月最后一天.
            Stack<int> sDate = new Stack<int>();
            Queue<AttendanceR> qAttendanceR = new Queue<AttendanceR>();
            //按指纹日期
            string fingerPrintDate = String.Empty;
            //导入数据的行数.
            int affectedCount = 0;
            //行最大值.
            int rowsMaxCount = 0;
            int colsMaxCount = 0;
            MyExcel myExcel = new MyExcel(xlsFilePath);
            //打开该文档。
            myExcel.open();
            #region 获取该工作簿中可视的工作表列表
            List<Worksheet> listVisualWS = myExcel.getVisualWS();
            List<AR_Sheet_Info> list_AR_Sheet_Info = new List<AR_Sheet_Info>();
            Usual_Excel_Helper uEHelper = null;
            foreach (Worksheet wS in listVisualWS)
            {
                int sheet_index = wS.Index;
                int max_rows = wS.UsedRange.Rows.Count;
                uEHelper = new Usual_Excel_Helper(wS);
                //获取A1内容
                string contentOfA3 = uEHelper.getSpecificCellValue("A3");
                //考勤时间为 空，忽略
                if (!"考勤时间".Equals(contentOfA3)) continue;
                int maxColumns = wS.UsedRange.Columns.Count;
                if (maxColumns <= 27 || maxColumns >= 32) continue;
                //列的宽度28-31.
                AR_Sheet_Info ar_sheet_info = new AR_Sheet_Info(sheet_index, max_rows);
                list_AR_Sheet_Info.Add(ar_sheet_info);
            }
            #endregion
            //获取list_AR_Sheet_Info 中 rows 最少的那个Sheet
            //先对list_AR_Sheet_Info进行排序。
            list_AR_Sheet_Info.Sort((a, b) => a.Max_rows.CompareTo(b.Max_rows));
            int sheetIndex = list_AR_Sheet_Info[0].SheetIndex;
            Worksheet ws = listVisualWS[sheetIndex - 1];
            //行;列最大值 赋值.
            rowsMaxCount = ws.UsedRange.Rows.Count;
            colsMaxCount = ws.UsedRange.Columns.Count;
            AttendanceR.Sheet_name = ws.Name;
            //判断首行是否为 考勤记录表;以此判断此表是否为考勤记录表.
            string A1Str = ((Range)ws.Cells[1, 1]).Text.ToString().Trim().Replace("\n", "").Replace("\r", "").Replace(" ", "");
            if (String.IsNullOrEmpty(A1Str))
            {
                msg.Msg = "工作表的A1单元格不能为空！";
                msg.Flag = false;
                myExcel.close();
                return msg;
            }
            //如果A1Str的内容不包含"考勤记录表"5个字。       
            if (!A1Str.Contains("考勤记录表"))
            {
                msg.Msg = "A1内容未包含'考勤记录表'";
                msg.Flag = false;
                myExcel.close();
                return msg;
            }
            #region 判断名称中是否区分了考勤记录。
            string Seq_Attendance_Record = string.Empty;
            int indexOfFullStop = xlsFilePath.LastIndexOf(".");
            Seq_Attendance_Record = xlsFilePath.Substring(indexOfFullStop - 1, 1);
            if (!CheckPattern.CheckNumber(Seq_Attendance_Record))
            {
                msg.Msg = "考勤记录表名称请以数字结尾！";
                msg.Flag = false;
                myExcel.close();
                return msg;
            }
            #endregion
            string excelName = Usual_Excel_Helper.getExcelName(xlsFilePath);
            AttendanceR.Prefix_Job_Number = excelName.Substring(excelName.Length - 1, 1).ToCharArray()[0];
            string C3Str = ((Range)ws.Cells[3, 3]).Text.ToString().Trim();
            //  \0: 表空字符.
            if (String.IsNullOrEmpty(C3Str))
            {
                msg.Msg = "异常: 考勤时间为空!";
                msg.Flag = false;
                myExcel.close();
                return msg;
            }
            //
            string[] ArrayC3 = C3Str.Split('~');
            if (ArrayC3.Length == 0)
            {
                msg.Msg = "异常: 考勤时间格式变更!";
                msg.Flag = false;
                myExcel.close();
                return msg;
            }
            AttendanceR.Start_date = ArrayC3[0].ToString().Trim();
            AttendanceR.End_date = ArrayC3[1].ToString().Trim();
            //制表时间:  L3 3行12列.
            string L3Str = ((Range)ws.Cells[3, 12]).Text.ToString().Trim();
            if (String.IsNullOrEmpty(L3Str))
            {
                msg.Msg = "异常: 制表时间为空!";
                msg.Flag = false;
                myExcel.close();
                return msg;
            }
            //制表时间.
            AttendanceR.Tabulation_time = L3Str;
            //检查第4行是否为;考勤时间:
            string A4Str = ((Range)ws.Cells[4, 1]).Text.ToString().Trim();
            if (!"1".Equals(A4Str, StringComparison.CurrentCultureIgnoreCase))
            {
                msg.Msg = "异常: 第四行已变更!";
                msg.Flag = false;
                myExcel.close();
                return msg;
            }
            uEHelper = new Usual_Excel_Helper(ws);
            //此刻不能删除，只是获取行号。
            Queue<Range> rangeToDelQueue = new Queue<Range>();
            //判断是否有空行。
            for (int i = 5; i <= rowsMaxCount; i++)
            {
                if (uEHelper.isBlankRow(i))
                {
                    //只要上一列不是
                    //删除掉此行。
                    //判断上一行中的A列是否为工号。
                    string temp = uEHelper.getSpecificCellValue("A" + (i - 1).ToString());
                    if ("工号:".Equals(temp))
                    {
                        continue;
                    }
                    //获取该行。
                    Range rangeToDel = (Microsoft.Office.Interop.Excel.Range)uEHelper.WS.Rows[i, System.Type.Missing];
                    //不为工号
                    rangeToDelQueue.Enqueue(rangeToDel);
                };
            }
            Range rangeToDelete;
            //开始删除空行。  
            while (rangeToDelQueue.Count > 0)
            {
                rangeToDelete = rangeToDelQueue.Dequeue();
                rangeToDelete.Delete(XlDeleteShiftDirection.xlShiftUp);
            };
            rowsMaxCount = ws.UsedRange.Rows.Count;
            //进度条长度增加。
            pbLength += colsMaxCount;
            pbLength += (colsMaxCount * (rowsMaxCount - 5 + 1));
            pI.Max_value = pbLength;
            pI.Current_value = 0;
            pI.Msg = excelName + "，正在读取：";
            pI.update();
            Console.WriteLine(excelName + "，正在读取：");
            //入队列值0
            sDate.Push(0);
            int actualMaxDay = 0;
            //开始循环
            for (int i = 1; i <= colsMaxCount; i++)
            {
                A4Str = ((Range)ws.Cells[4, i]).Text.ToString();
                //碰到第4行某列为空，退出循环。
                if (String.IsNullOrEmpty(A4Str))
                {
                    break;
                }
                int aDate = 0;
                //对A4Str进行分析.
                if (!Int32.TryParse(A4Str, out aDate))
                {
                    msg.Msg = String.Format(@"异常: 考勤日期行第{0}列出现非数字内容!", aDate);
                    msg.Flag = false;
                    myExcel.close();
                    return msg;
                }
                pI.Current_value++;
                pI.update();
                //判断新增的日期是否大于上一个.
                if (aDate <= sDate.Peek())
                {
                    //跳出循环.
                    break;
                }
                actualMaxDay++;
                sDate.Push(aDate);
            }
            //取其中的最小值。
            colsMaxCount = Math.Min(sDate.Count - 1, actualMaxDay);
            //考勤日期
            fingerPrintDate = AttendanceR.Start_date.Substring(0, 7).Replace('/', '-');
            string tempStr = string.Empty;
            //开始循环
            for (int colIndex = 1; colIndex <= colsMaxCount; colIndex++)
            {
                //从第5行开始.
                //奇数;偶数行共用一个对象.
                AttendanceR AR = null;
                for (int rowIndex = 5; rowIndex <= rowsMaxCount; rowIndex++)
                {
                    //如果行数为奇数则为工号行.
                    if (rowIndex % 2 == 1)
                    {
                        //工号行.
                        //取工号
                        AR = new AttendanceR();
                        AR.Job_number = ((Range)ws.Cells[rowIndex, 3]).Text.ToString().Trim();
                        //自行拼凑AR.
                        AR.combine_Job_Number();
                        //取姓名:  K5 
                        AR.Name = ((Range)ws.Cells[rowIndex, Usual_Excel_Helper.getColIndexByStr("K")]).Text.ToString().Trim();
                        //取部门: U5
                        AR.Dept = ((Range)ws.Cells[rowIndex, Usual_Excel_Helper.getColIndexByStr("U")]).Text.ToString().Trim();
                        //部门为空，则填充为NULL;
                        AR.Dept = !String.IsNullOrEmpty(AR.Dept) ? AR.Dept : "NULL";
                        //取日期.填充0;
                        AR.Fingerprint_Date = fingerPrintDate + "-" + colIndex.ToString().PadLeft(2, '0');
                    }
                    else
                    {
                        //偶数行取考勤结果.
                        //上班时间. 如B10;
                        tempStr = ((Range)ws.Cells[rowIndex, colIndex]).Text.ToString().Trim();
                        string tempFirstTime = String.Empty;
                        string tempLastTime = String.Empty;
                        if (!getFPTime(tempStr, out tempFirstTime, out tempLastTime))
                        {
                            msg.Msg = string.Format(@"导入失败：表中第{0}行{1}列的按指纹时间格式不对！", rowIndex, colIndex);
                            msg.Flag = false;
                            return msg;
                        };
                        AR.FPT_Fisrt_Time = String.IsNullOrEmpty(tempFirstTime) ? String.Empty : AR.Fingerprint_Date + " " + tempFirstTime;
                        AR.FPT_Last_Time = String.IsNullOrEmpty(tempLastTime) ? String.Empty : AR.Fingerprint_Date + " " + tempLastTime;
                        qAttendanceR.Enqueue(AR);
                    }
                    pI.Current_value++;
                    pI.update();
                }
            }
            Console.WriteLine("提交数据: ");
            pI.Max_value = qAttendanceR.Count;
            pI.Current_value = 0;
            pI.Msg = "提交数据: ";
            pI.update();
            System.Threading.Thread.Sleep(2000);
            #region
            OracleDaoHelper.noLogging("Attendance_Record");
            //保存对象
            while (qAttendanceR.Count > 0)
            {
                try
                {
                    affectedCount += qAttendanceR.Dequeue().saveDataToDB();
                    pI.Current_value++;
                    pI.update();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), "提示:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    msg.Msg = ex.ToString();
                    msg.Flag = false;
                    OracleConnHelper.closeConn();
                    return msg;
                }
            }
            #endregion
            OracleDaoHelper.logging("Attendance_Record");
            //释放对象
            myExcel.close();
            Console.WriteLine(String.Format(@"导入完成;计{0}条.", affectedCount.ToString()));
            //隐藏进度条。
            pI.Msg = String.Format(@"导入完成;计{0}条.", affectedCount.ToString());
            pI.Finish_flag = true;
            pI.update();
            return msg;
        }
        /// <summary>
        /// 获取第四行中，为天数的最大列索引
        /// </summary>
        /// <param name="wS"></param>
        /// <returns></returns>
        public static int getMaxColIndexOfThe4thRowOfAR(Worksheet wS) {
            Stack<int> sDate = new Stack<int>();
            sDate.Push(0);
            int aDate = 0;
            int maxColIndex = wS.UsedRange.Columns.Count;
            for (int colIndex = 1; colIndex <= maxColIndex; colIndex++)
            {
                Usual_Excel_Helper uEHelper = new Usual_Excel_Helper(wS);
                string tempStr = uEHelper.getCellContentByRowAndColIndex(4, colIndex);
                if (string.IsNullOrEmpty(tempStr))
                {
                    return colIndex - 1;
                }
                aDate = int.Parse(tempStr);
                //判断新增的日期是否大于上一个.
                if (aDate <= sDate.Peek())
                {
                    return colIndex - 1;
                }
                sDate.Push(aDate);
                //取其中的最小值。
            }
            return maxColIndex;
        }
        /// <summary>
        /// 判断某行是否都为数字。
        /// </summary>
        /// <param name="wS"></param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        public static bool isAllDigit(Worksheet wS, int rowIndex,out int checkedColIndex)
        {
            Usual_Excel_Helper uEHelper = new Usual_Excel_Helper(wS);
            int maxColIndex = wS.UsedRange.Columns.Count;
            bool flag = false;
            int num = 0;
            for (int colIndex = 1; colIndex <= maxColIndex; colIndex++)
            {
                string tempStr = uEHelper.getCellContentByRowAndColIndex(4, colIndex);
                flag = int.TryParse(tempStr, out num);
                if (!flag) {
                    checkedColIndex = colIndex;
                    return false;
                }
            }
            checkedColIndex = maxColIndex;
            return true;
        }
        /// <summary>
        /// 
        /// </summary>
        public static void clearSheet(Worksheet firstWS) {
            Queue<Range> rangeToDelQueue = new Queue<Range>();
            int rowsMaxCount;
            rowsMaxCount = firstWS.UsedRange.Rows.Count;
            Usual_Excel_Helper uEHelper = new Usual_Excel_Helper(firstWS);
            //获取最大列
            int maxColIndex = getMaxColIndexOfThe4thRowOfAR(firstWS);
            //判断是否有空行。
            for (int i = 5; i <= rowsMaxCount; i++)
            {
                if (uEHelper.isBlankRangeTheSpecificRow(i, 1, maxColIndex))
                {
                    //只要上一列不是
                    //删除掉此行。
                    //判断上一行中的A列是否为工号。
                    string temp = uEHelper.getSpecificCellValue("A" + (i - 1).ToString());
                    if ("工号:".Equals(temp))
                    {
                        continue;
                    }
                    //获取该行。
                    Range rangeToDel = (Microsoft.Office.Interop.Excel.Range)uEHelper.WS.Rows[i, System.Type.Missing];
                    //不为工号
                    rangeToDelQueue.Enqueue(rangeToDel);
                }
            }
            Range rangeToDelete;
            //开始删除空行。  
            while (rangeToDelQueue.Count > 0)
            {
                rangeToDelete = rangeToDelQueue.Dequeue();
                rangeToDelete.Delete(XlDeleteShiftDirection.xlShiftUp);
            }
        }
      
    }
}
