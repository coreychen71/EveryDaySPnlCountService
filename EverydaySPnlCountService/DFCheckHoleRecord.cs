using System;
using System.Data.SqlClient;
using System.Data;

namespace EverydaySPnlCountService
{
    class DFCheckHoleRecord
    {
        private string strCon = "server=EWNAS;database=ME;uid=me;pwd=2dae5na";

        /// <summary>
        /// 取得鑽孔生產申報紀錄
        /// </summary>
        /// <param name="WorkDate">查詢的日期</param>
        /// <returns></returns>
        public DataTable GetDFRecord(string WorkDate)
        {
            var result = new DataTable();
            var strComm = "select departname as '部門',empname as '人員',lotnum as '批號',partnum as '料號'," +
                "layername as '層別',workqnty as '數量',process as '工序',machineno as '機台',LAYER as '途程'," +
                "CONVERT(char(19), starttime, 120) as '開始時間',CONVERT(char(19), endtime, 120) as '結束時間'," +
                "rework as '重工' from drymcse " +
                "where departname = 'DF' and process = '鑽孔機' and workdate = '" + WorkDate + "' and " +
                "lotnum like 'M%' and todo=1 order by starttime desc";
            using (SqlConnection sqlcon = new SqlConnection(strCon))
            {
                using (SqlCommand sqlcomm = new SqlCommand(strComm, sqlcon))
                {
                    try
                    {
                        sqlcon.Open();
                        SqlDataReader read = sqlcomm.ExecuteReader();
                        result.Load(read);
                    }
                    catch (Exception ex)
                    {
                        MainMethod.InsertLog(ex.Message);
                    }
                }
            }
            return result;
        }
    }
}
