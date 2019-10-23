using AttendanceRecord.Entities;
using AttendanceRecord.Helper;
using Excel;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Tools;
using System.IO;
namespace AttendanceRecord
{
    public partial class FrmImportAR : Form
    {
        public FrmImportAR()
        {
            InitializeComponent();
        }
        string randomStr = String.Empty;
        public static string _action = "AttendanceRecord";
        AttendanceR aR = new AttendanceR();
        /// <summary>
        /// 用于存储过程中的参数。
        /// </summary>
        OracleHelper oH = OracleHelper.getBaseDao();
        string defaultDir = System.Windows.Forms.Application.StartupPath + @"\考勤记录";
        string xlsFilePath = String.Empty;

        MyExcel _1th_my_excel = null;
        MyExcel _2nd_my_excel = null;
        MyExcel _3rd_my_excel = null;
        MyExcel _4th_my_excel = null;
        Worksheet _1th_Sheet = null;
        Worksheet _2nd_Sheet = null;
        Worksheet _3rd_Sheet = null;
        Worksheet _4th_Sheet = null;

        Worksheet uncertainWS = null;
        string _defaultDir = System.Windows.Forms.Application.StartupPath + "\\uncertainRecord";
        string _uncertainWSPath = string.Empty;
        Range srcRange = null;
        Range destRange = null;
        int currentRow = 0;
        private void btnImportEmpsInfo_Click(object sender, EventArgs e)
        {
            btnViewTheUncertaiRecordInExcel.Enabled = false;
            lblResult.Text = "";
            lblResult.BackColor = this.BackColor;
            lblResult.Visible = false;
            //判断是否存在Excel进程.
            if (CmdHelper.ifExistsTheProcessByName("EXCEL"))
            {
                FrmPrompt frmPrompt = new FrmPrompt();
                frmPrompt.ShowDialog();
            }
            _uncertainWSPath = _defaultDir + "\\uncertainRecord_" + TimeHelper.getCurrentTimeStr() + ".xls";
            dgv.DataSource = null;
            lblResult.Visible = false;
            lblResult.Text = "";
            lblResult.BackColor = this.BackColor;
            tb.Clear();
            randomStr = TimeHelper.getCurrentTimeStr();
            xlsFilePath = FileNameDialog.getSelectedFilePathWithDefaultDir("请选择考勤记录：", "*.xls|*.xls", defaultDir);
            string dir = DirectoryHelper.getDirOfFile(xlsFilePath);
            if (string.IsNullOrEmpty(dir))
            {
                return;
            }
            List<string> xlsFileList = DirectoryHelper.getXlsFileUnderThePrescribedDir(dir);
            List<string> resultList = new List<string>();
            foreach (string xlsFile in xlsFileList)
            {
                string fileName = DirectoryHelper.getFileNameWithoutSuffix(xlsFile);
                if (!CheckString.CheckARName(fileName))
                {
                    continue;
                }
                //格式符合:  3月考勤记录1。
                resultList.Add(xlsFile);
            }
            #region 先判断第四行，是否全为数字。
            int maxColIndex = 0;
            if (!check4thRow(resultList, out maxColIndex))
            {
                return;
            }
            #endregion
            if (cbCheckSameNamesButDifferentMachineNo.Checked)
            {
                #region 保存关键信息到后台.
                saveCriticalARInfo(resultList);
                #endregion
                #region  打开4个考勤文件
                for (int i = 1; i <= resultList.Count; i++)
                {
                    switch (i)
                    {
                        case 1:
                            _1th_my_excel = new MyExcel(resultList[0]);
                            _1th_my_excel.open();
                            _1th_Sheet = _1th_my_excel.getFirstWorkSheetAfterOpen();
                            break;
                        case 2:
                            _2nd_my_excel = new MyExcel(resultList[1]);
                            _2nd_my_excel.open();
                            _2nd_Sheet = _2nd_my_excel.getFirstWorkSheetAfterOpen();
                            break;
                        case 3:
                            _3rd_my_excel = new MyExcel(resultList[2]);
                            _3rd_my_excel.open();
                            _3rd_Sheet = _3rd_my_excel.getFirstWorkSheetAfterOpen();
                            break;
                        case 4:
                            _4th_my_excel = new MyExcel(resultList[3]);
                            _4th_my_excel.open();
                            _4th_Sheet = _4th_my_excel.getFirstWorkSheetAfterOpen();
                            break;
                    }
                }
                #endregion
                #region 创建 _uncertain_myExcel;
                MyExcel uncertainRecordExcel = null;
                uncertainRecordExcel = new MyExcel(_uncertainWSPath);
                uncertainRecordExcel.create();
                uncertainRecordExcel.openWithoutAlerts();
                uncertainWS = uncertainRecordExcel.getFirstWorkSheetAfterOpen();
                //先写，日期行。
                Usual_Excel_Helper uEHelper = new Usual_Excel_Helper(uncertainWS);
                uEHelper.writeToSpecificRow(1, 1, maxColIndex);
                #endregion
                System.Data.DataTable dt = getSamePinYinButName();
                int amountOfGroupOfSamePinYinButName = getAmountOfGroupOfSamePinYinButName();
                bool have_same_pinyin_flag = false;
                if (dt != null && dt.Rows.Count > 0)
                {
                    have_same_pinyin_flag = true;
                }
                //*************判断是否拼音相同 开始********************8
                if (have_same_pinyin_flag)
                {

                    ShowResult.show(lblResult, "存在姓名拼音相同的记录!", false);
                    this.lblPrompt.Visible = false;
                    timerRestoreTheLblResult.Enabled = true;
                    #region 写记录到不确定文档中.
                    int theRowIndex = 0;
                    int Attendance_Machine_No = 0;
                    for (int i = 0; i <= dt.Rows.Count - 1; i++)
                    {
                        theRowIndex = int.Parse(dt.Rows[i]["行号"].ToString());
                        Attendance_Machine_No = int.Parse(dt.Rows[i]["卡机编号"].ToString());
                        switch (Attendance_Machine_No)
                        {
                            case 1:
                                //获取源区域
                                //替换源文件的工号为  工号位于第三列
                                _1th_Sheet.Cells[theRowIndex,3] = "'111111111" + ((Range)(_1th_Sheet.Cells[theRowIndex, 3])).Text.ToString().PadLeft(3,'0');
                                srcRange = _1th_Sheet.Range[_1th_Sheet.Cells[theRowIndex, 1], _1th_Sheet.Cells[theRowIndex + 1, maxColIndex]];
                                srcRange.Copy(Type.Missing);
                                //向目标复制。
                                //或取目标单元格。
                                currentRow = uncertainWS.UsedRange.Rows.Count;

                                destRange = uncertainWS.Range[uncertainWS.Cells[currentRow + 1, 1], uncertainWS.Cells[currentRow + 2, maxColIndex]];
                                //destRange.Select();
                                uncertainWS.Paste(destRange, false);
                                //保存一下。
                                break;
                            case 2:
                                _2nd_Sheet.Cells[theRowIndex, 3] = "'222222222" + ((Range)(_2nd_Sheet.Cells[theRowIndex, 3])).Text.ToString().PadLeft(3,'0');
                                srcRange = _2nd_Sheet.Range[_2nd_Sheet.Cells[theRowIndex, 1], _2nd_Sheet.Cells[theRowIndex + 1, maxColIndex]];
                                srcRange.Cells.Copy(Type.Missing);
                                //向目标复制。
                                //或取目标单元格。
                                currentRow = uncertainWS.UsedRange.Rows.Count;

                                destRange = uncertainWS.Range[uncertainWS.Cells[currentRow + 1, 1], uncertainWS.Cells[currentRow + 2, maxColIndex]];
                                //destRange.Select();
                                uncertainWS.Paste(destRange, false);
                                break;
                            case 3:
                                _3rd_Sheet.Cells[theRowIndex, 3] = "'333333333" + ((Range)(_3rd_Sheet.Cells[theRowIndex, 3])).Text.ToString().PadLeft(3,'0');
                                srcRange = _3rd_Sheet.Range[_3rd_Sheet.Cells[theRowIndex, 1], _3rd_Sheet.Cells[theRowIndex + 1, maxColIndex]];
                                srcRange.Cells.Copy(Type.Missing);
                                //向目标复制。
                                //或取目标单元格。
                                currentRow = uncertainWS.UsedRange.Rows.Count;
                                destRange = uncertainWS.Range[uncertainWS.Cells[currentRow + 1, 1], uncertainWS.Cells[currentRow + 2, maxColIndex]];
                                //destRange.Select();
                                uncertainWS.Paste(destRange, false);
                                break;
                            case 4:
                                _4th_Sheet.Cells[theRowIndex,3] = "'444444444" + ((Range)(_4th_Sheet.Cells[theRowIndex, 3])).Text.ToString().PadLeft(3,'0');
                                srcRange = _4th_Sheet.Range[_4th_Sheet.Cells[theRowIndex, 1], _4th_Sheet.Cells[theRowIndex + 1, maxColIndex]];
                                srcRange.Cells.Copy(Type.Missing);
                                //向目标复制。
                                //或取目标单元格。
                                currentRow = uncertainWS.UsedRange.Rows.Count;
                                destRange = uncertainWS.Range[uncertainWS.Cells[currentRow + 1, 1], uncertainWS.Cells[currentRow + 2, maxColIndex]];
                                //destRange.Select();
                                uncertainWS.Paste(destRange, false);
                                break;
                        }
                    }
                    //设置列宽
                    uncertainWS.UsedRange.ColumnWidth = 3.75;
                    //显示该文档对应的图片
                    #endregion
                    string promptStr = string.Format(@" 存在姓名拼音相同但书写不同的记录：{1}组;{0}
确定: 将视为不同员工;   取消: 取消本次导入;", "\r\n", amountOfGroupOfSamePinYinButName);
                    if (DialogResult.Cancel.Equals(MessageBox.Show(promptStr, "提示：", MessageBoxButtons.OKCancel, MessageBoxIcon.Information)))
                    {
                        closeThe4ARExcels();
                        uncertainWS.UsedRange.ColumnWidth = 3.75M;
                        uncertainRecordExcel.saveWithoutAutoFit();
                        
                        uncertainRecordExcel.close();
                        //显示该文档。
                        uncertainRecordExcel = new MyExcel(_uncertainWSPath);
                        uncertainRecordExcel.open(true);
                        btnViewTheUncertaiRecordInExcel.Enabled = true;
                        return;
                    }
                    if(!btnViewTheUncertaiRecordInExcel.Enabled) btnViewTheUncertaiRecordInExcel.Enabled = true; 
                }
              
                //*************判断是否拼音相同  结束*****************88
                //1.h
                dt = getSameNameInfo();
                //获取汉字相同的组的数目。
                int amountOfGroupOfSameName = getAmountOfGroupOfSameName();
                string prompt = string.Empty;
                if (dt.Rows.Count != 0)
                {
                    int theRowIndex = 0;
                    int Attendance_Machine_No = 0;
                    #region 同名记录书写结束.
                    for (int i = 0; i <= dt.Rows.Count - 1; i++)
                    {
                        theRowIndex = int.Parse(dt.Rows[i]["行号"].ToString());
                        Attendance_Machine_No = int.Parse(dt.Rows[i]["卡机编号"].ToString());
                        switch (Attendance_Machine_No)
                        {
                            case 1:
                                _1th_Sheet.Cells[theRowIndex, 3] = "'111111111" + ((Range)(_1th_Sheet.Cells[theRowIndex, 3])).Text.ToString().PadLeft(3, '0');
                                //获取源区域
                                srcRange = _1th_Sheet.Range[_1th_Sheet.Cells[theRowIndex, 1], _1th_Sheet.Cells[theRowIndex + 1, maxColIndex]];
                                srcRange.Copy(Type.Missing);
                                //向目标复制。
                                //或取目标单元格。
                                currentRow = uncertainWS.UsedRange.Rows.Count;
                                destRange = uncertainWS.Range[uncertainWS.Cells[currentRow + 1, 1], uncertainWS.Cells[currentRow + 2, maxColIndex]];
                                //destRange.Select();
                                uncertainWS.Paste(destRange, false);
                                //保存一下。
                                break;
                            case 2:
                                _2nd_Sheet.Cells[theRowIndex, 3] = "'222222222" + ((Range)(_2nd_Sheet.Cells[theRowIndex, 3])).Text.ToString().PadLeft(3, '0');
                                srcRange = _2nd_Sheet.Range[_2nd_Sheet.Cells[theRowIndex, 1], _2nd_Sheet.Cells[theRowIndex + 1, maxColIndex]];
                                srcRange.Cells.Copy(Type.Missing);
                                //向目标复制。
                                //或取目标单元格。
                                currentRow = uncertainWS.UsedRange.Rows.Count;
                                destRange = uncertainWS.Range[uncertainWS.Cells[currentRow + 1, 1], uncertainWS.Cells[currentRow + 2, maxColIndex]];
                                //destRange.Select();
                                uncertainWS.Paste(destRange, false);
                                break;
                            case 3:
                                _3rd_Sheet.Cells[theRowIndex, 3] = "'333333333" + ((Range)(_3rd_Sheet.Cells[theRowIndex, 3])).Text.ToString().PadLeft(3, '0');
                                srcRange = _3rd_Sheet.Range[_3rd_Sheet.Cells[theRowIndex, 1], _3rd_Sheet.Cells[theRowIndex + 1, maxColIndex]];
                                srcRange.Cells.Copy(Type.Missing);
                                //向目标复制。
                                //或取目标单元格。
                                currentRow = uncertainWS.UsedRange.Rows.Count;
                                destRange = uncertainWS.Range[uncertainWS.Cells[currentRow + 1, 1], uncertainWS.Cells[currentRow + 2, maxColIndex]];
                                //destRange.Select();
                                uncertainWS.Paste(destRange, false);
                                break;
                            case 4:
                                _4th_Sheet.Cells[theRowIndex, 3] = "'444444444" + ((Range)(_4th_Sheet.Cells[theRowIndex, 3])).Text.ToString().PadLeft(3, '0');
                                srcRange = _4th_Sheet.Range[_4th_Sheet.Cells[theRowIndex, 1], _4th_Sheet.Cells[theRowIndex + 1, maxColIndex]];
                                srcRange.Cells.Copy(Type.Missing);
                                //向目标复制。
                                //或取目标单元格。
                                currentRow = uncertainWS.UsedRange.Rows.Count;
                                destRange = uncertainWS.Range[uncertainWS.Cells[currentRow + 1, 1], uncertainWS.Cells[currentRow + 2, maxColIndex]];
                                //destRange.Select();
                                uncertainWS.Paste(destRange, false);
                                break;
                        }
                    }
                    #endregion
                    prompt = string.Format(@"  存在同名的记录：{1}组;{0}
确定: 将视为同一员工;   取消: 取消本次导入;", "\r\n",amountOfGroupOfSameName);
                    if (DialogResult.Cancel.Equals(MessageBox.Show(prompt, "提示：", MessageBoxButtons.OKCancel, MessageBoxIcon.Information)))
                    {
                        uEHelper.setAllColumnsWidth(3.75M);
                        uncertainRecordExcel.saveWithoutAutoFit();
                        uncertainRecordExcel.close();
                        //显示该文档。
                        uncertainRecordExcel = new MyExcel(_uncertainWSPath);
                        uncertainRecordExcel.open(true);
                        closeThe4ARExcels();
                        if (!btnViewTheUncertaiRecordInExcel.Enabled) btnViewTheUncertaiRecordInExcel.Enabled = true;
                        return;
                    }
                    if (!btnViewTheUncertaiRecordInExcel.Enabled) btnViewTheUncertaiRecordInExcel.Enabled = true;
                }
                //关闭不确定文档。
                uEHelper.setAllColumnsWidth(3.75M);
                uncertainRecordExcel.saveWithoutAutoFit();
                uncertainRecordExcel.close();
                closeThe4ARExcels();
            }
            resultList.Sort();
            //判断该考勤机中是否已经存在，某月的记录
            foreach (string xlsFilePath in resultList)
            {
                string fileName = DirectoryHelper.getFileNameWithoutSuffix(xlsFilePath);
                int attendanceMachineFlag = int.Parse(fileName.Substring(fileName.Length - 1, 1));
                //打开文档获取考勤机，所记录的日期。
                MyExcel myExcel = new MyExcel(xlsFilePath);
                myExcel.open();
                Worksheet firstSheet = myExcel.getFirstWorkSheetAfterOpen();
                Usual_Excel_Helper uEHelper = new Usual_Excel_Helper(firstSheet);
                string year_and_month_str = uEHelper.getCellContentByRowAndColIndex(3, 3);
                year_and_month_str = year_and_month_str.Substring(0, 7);
                myExcel.close();
                System.Data.DataTable dtSubmitInfo = getSubmitInfoOfTheSpecificeMachineAndYearAndMonth(attendanceMachineFlag, year_and_month_str);
                if (!"0".Equals(dtSubmitInfo.Rows[0]["nums_of_staffs"].ToString()))
                {
                    string prompt = string.Format(@"    您已经于: {1},{0}
提交了{2} 考勤机{3},{0}
共计{4}个用户的纪录。{0}    
确定：覆盖；取消：退出?", "\r\n", dtSubmitInfo.Rows[0]["latest_record_time"].ToString(), year_and_month_str, attendanceMachineFlag, dtSubmitInfo.Rows[0]["nums_of_staffs"].ToString());
                    if (DialogResult.Cancel.Equals(MessageBox.Show(prompt, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)))
                    {
                        return;
                    }
                    delTheInfoOfTheSpecificeMachineAndYearAndMonth(attendanceMachineFlag, year_and_month_str);
                }
            }
            this.dgv.DataSource = null;
            //this.dgv.Columns.Clear();
            lblPrompt.Visible = false;
            lblPrompt.Text = "";
            pb.Value = 0;
            pb.Maximum = 0;
            pb.Visible = false;
            foreach (string xlsFilePath in resultList)
            {
                tb.Text = xlsFilePath;
                lblResult.Visible = false;
                MSG msg = AttendanceRHelper.ImportAttendanceRecordToDB(xlsFilePath, randomStr, lblPrompt, pb, lblResult);
                //导入完成后进行保存，保存该文件至prepared目录中
                pb.Visible = false;
                lblPrompt.Visible = false;
                ShowResult.show(lblResult, msg.Msg, msg.Flag);
                //timerRestoreTheLblResult.Enabled = true;
                if (!msg.Flag) return;
                //saveTheExcel(xlsFilePath);
            }
            //加载导入的数据。
            this.dgv.DataSource = null;
            this.dgv.DataSource = aR.getAR(randomStr);
            this.dgv.Visible = true;
            DGVHelper.AutoSizeForDGV(dgv);
            tb.Clear();
        }
        private void closeThe4ARExcels()
        {
            //检查结束.
            #region 关闭4个考勤文件
            if (_1th_my_excel != null)
            {
                _1th_my_excel.close();
            }
            if (_2nd_my_excel != null)
            {
                _2nd_my_excel.close();
            }
            if (_3rd_my_excel != null)
            {
                _3rd_my_excel.close();
            }
            if (_4th_my_excel != null)
            {
                _4th_my_excel.close();
            }
            #endregion
        }
        private bool check4thRow(List<String> excelPathList, out int maxColIndex)
        {
            maxColIndex = 0;
            //先清除所有记录。
            AR_Temp.deleteTheARTemp();
            foreach (string excelPath in excelPathList)
            {
                //打开文档
                MyExcel myExcel = new MyExcel(excelPath);
                myExcel.open();
                Worksheet firstWS = myExcel.getFirstWorkSheetAfterOpen();
                string fileNameWithoutSuffix = DirectoryHelper.getFileNameWithoutSuffix(excelPath);
                int checkedColIndex = 0;
                if (!AttendanceRHelper.isAllDigit(firstWS, 4, out checkedColIndex))
                {
                    myExcel.close();
                    lblPrompt.Visible = false;
                    ShowResult.show(lblResult, fileNameWithoutSuffix + ": 第4行" + checkedColIndex.ToString() + "列非数字;   导入取消。", false);
                    //timerRestoreTheLblResult.Start();
                    return false;
                }
                if (maxColIndex == 0)
                {
                    Usual_Excel_Helper uEHelper = new Usual_Excel_Helper(firstWS);
                    maxColIndex = uEHelper.getMaxColIndexBeforeBlankCellInSepcificRow(4);
                }
                myExcel.close();
            }
            return true;
        }
        private System.Data.DataTable getSubmitInfoOfTheSpecificeMachineAndYearAndMonth(int attendanceMachineFlag, string year_and_month_str)
        {
            string sqlStr = string.Format(@"Select count(distinct job_number) as nums_of_staffs,
                                                    max(to_char(record_time,'yyyy-MM-dd HH24:MI:SS')) as latest_record_time
                                             from Attendance_Record
                                             where substr(job_number,1,1) = '{0}'
                                             and trunc(fingerprint_date,'MM') = to_date('{1}','yyyy-MM')",
                                             attendanceMachineFlag,
                                             year_and_month_str);
            return OracleDaoHelper.getDTBySql(sqlStr);
        }
        private void delTheInfoOfTheSpecificeMachineAndYearAndMonth(int attendanceMachineFlag, string year_and_month_str)
        {
            string sqlStr = string.Format(@"delete 
                                            from Attendance_record 
                                            where substr(job_number,1,1) = '{0}'
                                            and trunc(start_date,'MM') = to_date('{1}','yyyy-MM')",
                                            attendanceMachineFlag,
                                            year_and_month_str);
            OracleDaoHelper.executeSQL(sqlStr);
        }
        private System.Data.DataTable getSameNameInfo()
        {
            string sqlStr = string.Format(@"select distinct 
                                                            AR_Temp.Job_Number AS ""工号"",
                                                            AR_Temp.name AS ""姓名"",
                                                            AR_Temp.Attendance_Machine_Flag AS ""卡机编号"",
                                                            AR_Temp.Row_Index AS ""行号"",
                                                            AR_Temp.Record_Time  AS ""记录时间""
                                            from AR_Temp, (
                                                           select job_number,
                                                                    name,
                                                                    attendance_machine_flag,
                                                                    row_index,
                                                                    record_time
                                                           from AR_Temp
                                             ) AR_T
                                            WHERE AR_Temp.name = AR_T.Name
                                            AND(
                                                  (AR_Temp.Attendance_Machine_Flag = AR_T.attendance_machine_flag
                                                  AND AR_Temp.Job_Number != AR_T.job_number)
                                                  OR(
                                                     AR_Temp.Attendance_Machine_Flag != AR_T.attendance_machine_flag
                                                  )
                                            )
                                            order by NLSSORT(""姓名"", 'NLS_SORT= SCHINESE_PINYIN_M')");
            System.Data.DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            return dt;
        }
        /// <summary>
        /// 获取汉字相同的记录的组数。
        /// </summary>
        /// <returns></returns>
        private int getAmountOfGroupOfSameName() {
            string sqlStr = string.Format(@"SELECT ""姓名""
                                            FROM
                                            (
                                                select distinct
                                                                AR_Temp.Job_Number AS ""工号"",
                                                                AR_Temp.name AS ""姓名"",
                                                                AR_Temp.Attendance_Machine_Flag AS ""卡机编号"",
                                                                AR_Temp.Row_Index AS ""行号"",
                                                                AR_Temp.Record_Time  AS ""记录时间""
                                                from AR_Temp, (
                                                               select job_number,
                                                                        name,
                                                                        attendance_machine_flag,
                                                                        row_index,
                                                                        record_time
                                                               from AR_Temp
                                                 ) AR_T
                                                WHERE AR_Temp.name = AR_T.Name
                                                AND(
                                                      (AR_Temp.Attendance_Machine_Flag = AR_T.attendance_machine_flag
                                                      AND AR_Temp.Job_Number != AR_T.job_number)
                                                      OR(
                                                         AR_Temp.Attendance_Machine_Flag != AR_T.attendance_machine_flag
                                                      )
                                                )
                                            )   T
                                            group by ""姓名""");
            return OracleDaoHelper.getDTBySql(sqlStr).Rows.Count;
        }
        private System.Data.DataTable getSamePinYinButName()
        {
            string sqlStr = string.Format(@"select distinct 
                                                                AR_Temp.Job_Number AS ""工号"",
                                                                AR_Temp.name AS ""姓名"",
                                                                AR_Temp.Attendance_Machine_Flag AS ""卡机编号"",
                                                                AR_Temp.Row_Index AS ""行号"",
                                                                AR_Temp.Record_Time  AS ""记录时间""
                                                from AR_Temp, (
                                                              select distinct name
                                                              from AR_Temp
                                                ) AR_T
                                                WHERE AR_Temp.name ! = AR_T.Name
                                                AND gethzpy.GetHzFullPY(AR_Temp.name) = gethzpy.GetHzFullPY(AR_T.name)
                                                order by NLSSORT(""姓名"", 'NLS_SORT= SCHINESE_PINYIN_M')");
            return OracleDaoHelper.getDTBySql(sqlStr);
        }
        /// <summary> 
        /// 获取姓名拼音相同，但汉字书写不同的记录的组数。
        /// </summary>
        /// <returns></returns>
        private int getAmountOfGroupOfSamePinYinButName() {
            string sqlStr = string.Format(@"select 1
                                                    from 
                                                    (
                                                        select distinct 
                                                                        AR_Temp.Job_Number AS ""工号"",
                                                                        AR_Temp.name AS ""姓名"",
                                                                        AR_Temp.Attendance_Machine_Flag AS ""卡机编号"",
                                                                        AR_Temp.Row_Index AS ""行号"",
                                                                        AR_Temp.Record_Time  AS ""记录时间""
                                                        from AR_Temp, (
                                                                      select distinct name
                                                                      from AR_Temp
                                                        ) AR_T
                                                        WHERE AR_Temp.name ! = AR_T.Name
                                                        AND gethzpy.GetHzFullPY(AR_Temp.name) = gethzpy.GetHzFullPY(AR_T.name)
                                                    ) T
                                                    group by gethzpy.GetHzFullPY(""姓名"")");
            return OracleDaoHelper.getDTBySql(sqlStr).Rows.Count;
        }
        private void saveCriticalARInfo(List<String> excelPathList)
        {
            //先清除所有记录。
            AR_Temp.deleteTheARTemp();
            foreach (string excelPath in excelPathList)
            {
                //打开文档
                MyExcel myExcel = new MyExcel(excelPath);
                myExcel.open();
                Worksheet firstWS = myExcel.getFirstWorkSheetAfterOpen();
                //删除  时间后立即为空的行。
                AttendanceRHelper.clearSheet(firstWS);
                Usual_Excel_Helper uEHelper = new Usual_Excel_Helper(firstWS);
                string fileNameWithoutSuffix = DirectoryHelper.getFileNameWithoutSuffix(excelPath);
                //先获取第4行的最大行列数目。
                int rowMaxIndex = firstWS.UsedRange.Rows.Count;
                pb.Value = 0;
                pb.Maximum = rowMaxIndex - 4;
                pb.Visible = true;
                lblPrompt.Visible = true;
                lblResult.Visible = false;
                lblPrompt.Text = fileNameWithoutSuffix + ": 基本信息采集中...";
                for (int rowIndex = 5; rowIndex <= rowMaxIndex; rowIndex++)
                {
                    if (0 == rowIndex % 2)
                    {
                        pb.Value++;
                        continue;
                    }
                    //姓名 存于第11列。
                    string name = uEHelper.getCellContentByRowAndColIndex(rowIndex, 11);
                    AR_Temp ar_Temp = new AR_Temp();
                    ar_Temp.Attendance_machine_flag = int.Parse(fileNameWithoutSuffix.Substring(fileNameWithoutSuffix.Length - 1, 1));
                    ar_Temp.Row_Index = rowIndex;
                    ar_Temp.Job_number = uEHelper.getCellContentByRowAndColIndex(rowIndex, 3);
                    ar_Temp.Name = name;
                    ar_Temp.saveRecord();
                    pb.Value++;
                }
                lblPrompt.Visible = false;
                pb.Visible = false;
                myExcel.save();
                myExcel.close();
            }
        }
        private void timerRestoreTheLblResult_Tick(object sender, EventArgs e)
        {
            timerRestoreTheLblResult.Enabled = false;
            lblResult.Text = "";
            lblResult.BackColor = this.BackColor;
            lblResult.Visible = false;
        }

        private void btnViewTheUncertaiRecordInExcel_Click(object sender, EventArgs e)
        {
            MyExcel uncertainRecordExcel = new MyExcel(_uncertainWSPath);
            uncertainRecordExcel.open(true);
        }
    }
}
