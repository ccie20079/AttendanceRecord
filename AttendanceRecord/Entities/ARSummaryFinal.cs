using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Oracle.DataAccess.Client;
using System.Data;
using Tools;
namespace AttendanceRecord.Entities
{
    public class ARSummaryFinal
    {
        public static DataTable getARSummaryFinal(string v_year_and_month_str ) {
            string procName = "PKG_AR_SUMMARY.getARSummary";
            OracleParameter param_year_and_month = new OracleParameter("v_year_and_month_str",OracleDbType.Varchar2,ParameterDirection.Input);
            OracleParameter param_cur_result = new OracleParameter("v_cur_result", OracleDbType.RefCursor, ParameterDirection.ReturnValue);
            param_year_and_month.Value = v_year_and_month_str;
            OracleParameter[] parameters = new OracleParameter[2] { param_cur_result,param_year_and_month  };
            OracleHelper oH = OracleHelper.getBaseDao();
            return oH.getDT(procName, parameters); 
       }
    }
}
