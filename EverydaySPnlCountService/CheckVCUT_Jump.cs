using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Data.SqlClient;

namespace EverydaySPnlCountService
{
    /// <summary>
    /// 檢查製程現帳中是否有跳刀料號且途程已達V-CUT前二站或已進V-CUT
    /// </summary>
    class CheckVCUT_Jump
    {
        private string strCon = "server=ERP;database=EW;uid=jsis;pwd=JSIS";

        /// <summary>
        /// 取得檢查結果
        /// </summary>
        /// <returns></returns>
        public DataTable GetResult()
        {
            var Result = new DataTable();
            using (SqlConnection sqlcon = new SqlConnection(strCon))
            {
                using (SqlCommand sqlcomm = new SqlCommand(string.Empty, sqlcon))
                {
                    sqlcomm.CommandType = CommandType.StoredProcedure;
                    sqlcomm.CommandText = "CheckVCUT_Jump";
                    try
                    {
                        sqlcon.Open();
                        SqlDataReader reader = sqlcomm.ExecuteReader();
                        Result.Load(reader);
                    }
                    catch (Exception ex)
                    {
                        MainMethod.InsertLog(ex.Message);
                    }
                }
            }
            return Result;
        }
    }
}
