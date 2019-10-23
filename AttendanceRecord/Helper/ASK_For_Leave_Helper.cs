﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Tools;
using Oracle.DataAccess.Client;
using System.Data;
namespace AttendanceRecord.Helper
{
    /// <summary>
    /// 请假帮助类。
    /// </summary>
    public class ASK_For_Leave_Helper
    {
        private string _name;
        private string _startTime;
        private string _endTime;
        private string _NO;     //业务单据号.

        public ASK_For_Leave_Helper(string name, string startTime, string endTime,string NO)
        {
            this._name = name;
            this._startTime = startTime;
            this._endTime = endTime;
            this._NO = NO;
        }
        public void save() {
            //先删除
            delTheTemp();
            string sqlStr = String.Format(@"insert into ask_for_leave_temp(
                                            job_number,
                                            name, 
                                            leave_start_time, 
                                            leave_end_time, 
                                            record_time, 
                                            type,
                                            no
                                      )
                                      select 
                                             job_number,
                                             name,
                                             to_date('{1}','yyyy-MM-dd hh24:mi:ss'),
                                             to_date('{2}','yyyy-MM-dd hh24:mi:ss'),
                                             sysdate,
                                             '{3}',
                                             '{4}'
                                     from employees
                                     where name = '{0}'",
                                     this._name,
                                     this._startTime,
                                     this._endTime,
                                     "",
                                     this._NO);
            Tools.OracleDaoHelper.executeSQL(sqlStr);
            delTheNO();
            saveTempToFormal();
        }
        #region 删除临时表中的记录.
        private void delTheTemp() {
            string sqlStr = String.Format("delete from ask_for_leave_temp");
            Tools.OracleDaoHelper.executeSQL(sqlStr);
        }
        #endregion
        #region 删除此单据
        private int  delTheNO() {
            string sqlStr = String.Format("delete from ask_for_leave where no = '{0}'",this._NO);
            return Tools.OracleDaoHelper.executeSQL(sqlStr);
        }
        #endregion;
        public static int delTheNO(string no) {
            string sqlStr = String.Format("delete from ask_for_leave where no = '{0}'", no);
            return Tools.OracleDaoHelper.executeSQL(sqlStr);
        }
        #region 将临时表中的数据，拆分后放入正式表中。
        private int saveTempToFormal() {
            string procedueName = "SaveToAFL";
            OracleHelper oH = OracleHelper.getBaseDao();
            OracleParameter[] parameters = { };
            return oH.ExecuteNonQuery(procedueName, parameters);
        }
        #endregion
        #region 获取所有请假条信息
        public static DataTable getAllVacationList() {
            string sqlStr = String.Format(@"select 
                                                emps.job_number AS ""工号"", 
                                                emps.name AS ""姓名"",
                                                a_f_l.leave_start_time AS ""起始时间"", 
                                                a_f_l.leave_end_time AS ""终止时间"", 
                                                a_f_l.record_time AS ""记录时间"", 
                                                a_f_l.no AS ""单号""
                                            from ask_for_leave a_f_l 
                                                inner join employees emps
                                                on a_f_l.job_number = emps.job_number
                                            order by a_f_l.seq desc");
            return OracleDaoHelper.getDTBySql(sqlStr);
        }
        #endregion
        #region 获取请假条的单号。
        public static string getLastedNO() {
            string sqlStr = String.Format(@"select 
                                            TO_CHAR(SYSDATE,'YYYYMMDD')||
                                            CASE 
                                                 WHEN NOT EXISTS(
                                                                 SELECT 1 
                                                                 FROM  Ask_For_Leave 
                                                                 where substr(NO,1,8) = TO_CHAR(SYSDATE,'YYYYMMDD')
                                                                 ) THEN '001'
                                                 ELSE ( select LPAD(MAX(SUBSTR(A_F_L.NO,9,3)) +1,3,'000')
                                                        FROM  Ask_For_Leave A_F_L
                                                        where substr(A_F_L.NO,1,8) = TO_CHAR(SYSDATE,'YYYYMMDD')
                                                        )
                                            END
                                    FROM DUAL");
            DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            return dt.Rows[0][0].ToString();
        }
        #endregion
        #region 判断此时间范围是否存在请假条
        public bool ifExistsAtRange() {
            string sqlStr = String.Format(@"
                                            SELECT 1 
                                            FROM ASK_FOR_LEAVE A_F_L
                                            WHERE A_F_L.name = '{0}' 
                                            AND TRUNC(TO_DATE('{1}','yyyy-MM-dd hh24:mi:ss'),'DD') = TRUNC(Leave_start_time) 
                                            ",
                                            this._name,
                                            this._startTime);
            return OracleDaoHelper.getDTBySql(sqlStr).Rows.Count > 0 ? true : false;
        }
        #endregion
        #region 判断此请假范围是否有休息日.人数小于99人,即被认为是休息日.
        public bool ifExistsVacationAtRange() {
            bool result = false;
            string sqlStr = String.Format(@"select 1
                                              from Attendance_Record AR
                                              where (AR.Fpt_First_Time IS NOT NULL OR AR.Fpt_Last_Time IS NOT NULL)
                                              AND trunc(AR.fingerprint_date,'DD') <= to_date('{0}','yyyy-MM-dd')
                                              and trunc(AR.fingerprint_date,'DD') >= to_date('{1}','yyyy-MM-dd')
                                              group by AR.Fingerprint_Date
                                              having count(1) < 99",
                                              this._startTime.Substring(0,10),
                                              this._endTime.Substring(0,10)
                                              );
            result = OracleDaoHelper.getDTBySql(sqlStr).Rows.Count > 0 ? true : false;
            return result;
        }
        #endregion
    }
}
