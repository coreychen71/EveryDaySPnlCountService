using System;
using System.Data;
using System.Data.SqlClient;

namespace EverydaySPnlCountService
{
    class ConnEWNAS
    {
        static string strCon = "server=EWNAS;database=ME;uid=me;pwd=2dae5na";

        /// <summary>
        /// 取得Ewproject每日未結案客訴事項(未逾期)
        /// </summary>
        /// <returns></returns>
        public DataTable EveryDayCustomerComplaint()
        {
            var result = new DataTable();
            var strComm = "select custermerid '客戶別',CONVERT(char(10),cpldate,120) '客訴日期',datecode '週期'," +
                "PARTNUM '料號',custermerpartnum '客戶料號',QNTY '數量',defectname '原因',discribe '描述'," +
                "result '處理結果',dutydp '責任單位',dutyman '責任人',CONVERT(char(10), requirdate, 120) '預計完成日'," +
                "writeman '登記人員',CONVERT(char(10), writedate, 120) '登記日期' from qacustermercpl " +
                "where stateflag is null order by cpldate desc,custermerid,PARTNUM";
            using (SqlConnection sqlcon = new SqlConnection(strCon))
            {
                SqlCommand sqlcomm = new SqlCommand(strComm, sqlcon);
                try
                {
                    sqlcon.Open();
                    SqlDataReader reader = sqlcomm.ExecuteReader();
                    result.Load(reader);
                }
                catch (Exception ex)
                {
                    MainMethod.InsertLog(ex.Message);
                }
            }
            return result;
        }

        /// <summary>
        /// 取得預報廢帳Log
        /// </summary>
        /// <returns></returns>
        public DataTable GetScrapWIPLog()
        {
            var result = new DataTable();
            var strComm = "select * from ScrapWIPLog";
            using (SqlConnection sqlcon = new SqlConnection(strCon))
            {
                SqlCommand sqlcomm = new SqlCommand(strComm, sqlcon);
                try
                {
                    sqlcon.Open();
                    SqlDataReader reader = sqlcomm.ExecuteReader();
                    result.Load(reader);
                }
                catch (Exception ex)
                {
                    MainMethod.InsertLog("ConnEWNAS.GetScrapWIPLog()-" + ex.Message);
                }
            }
            return result;
        }

        /// <summary>
        /// 寫入有下新製令的預報廢帳LOG
        /// </summary>
        /// <param name="IssuePaperNum">製令單號</param>
        /// <param name="IssuePaperDate">單據日期</param>
        /// <param name="IssuePartNum">料號</param>
        /// <param name="IssueRevision">版序</param>
        /// <param name="IssuePOTypeName">訂單類型</param>
        /// <param name="IssueQnty">製令單總數量</param>
        /// <param name="IssueExpStkTime">期望繳庫日</param>
        /// <param name="IssueFinishUser">審核人員</param>
        /// <param name="ScrapProcName">提出製程</param>
        /// <param name="ScrapProcCode">代碼</param>
        /// <param name="ScrapLotNum">批號</param>
        /// <param name="ScrapLotStatusName">狀態</param>
        /// <param name="ScrapLayerName">階段名稱</param>
        /// <param name="ScrapPOPName">形狀</param>
        /// <param name="ScrapPOP_SP">排版</param>
        /// <param name="ScrapQnty">預報廢數量</param>
        /// <param name="ScrapNotes">批量種類</param>
        /// <returns>int</returns>
        public int InsertScrapWIPLog(string IssuePaperNum, string IssuePaperDate, string IssuePartNum,
            string IssueRevision, string IssuePOTypeName, string IssueQnty, string IssueExpStkTime,
            string IssueFinishUser, string ScrapProcName, string ScrapProcCode, string ScrapLotNum,
            string ScrapLotStatusName, string ScrapLayerName, string ScrapPOPName, string ScrapPOP_SP,
            string ScrapQnty, string ScrapNotes)
        {
            var result = 0;
            var strComm = "insert into ScrapWIPLog(IssuePaperNum,IssuePaperDate,IssuePartNum,IssueRevision," +
                "IssuePOTypeName,IssueQnty,IssueExpStkTime,IssueFinishUser,ScrapProcName,ScrapProcCode,ScrapLotNum," +
                "ScrapLotStatusName,ScrapLayerName,ScrapPOPName,ScrapPOP_SP,ScrapQnty,ScrapNotes) " +
                "values('" + IssuePaperNum + "', '" + IssuePaperDate + "', '" + IssuePartNum + "', '" + IssueRevision +
                "', '" + IssuePOTypeName + "', '" + IssueQnty + "','" + IssueExpStkTime + "','" + IssueFinishUser +
                "','" + ScrapProcName + "','" + ScrapProcCode + "','" + ScrapLotNum + "','" + ScrapLotStatusName +
                "','" + ScrapLayerName + "','" + ScrapPOPName + "','" + ScrapPOP_SP + "','" + ScrapQnty + "','" +
                ScrapNotes + "')";
            using (SqlConnection sqlcon = new SqlConnection(strCon))
            {
                SqlCommand sqlcomm = new SqlCommand(strComm, sqlcon);
                try
                {
                    sqlcon.Open();
                    result = sqlcomm.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MainMethod.InsertLog("ConnEWNAS.InsertScrapWIPLog()-" + ex.Message);
                }
            }
            return result;
        }

        /// <summary>
        /// 刪除預報廢帳LOG
        /// </summary>
        /// <param name="ID">ID</param>
        /// <returns></returns>
        public int DeleteScrapWIPLog(string ID)
        {
            var result = 0;
            var strComm = "delete ScrapWIPLog where ID=" + ID;
            using (SqlConnection sqlcon = new SqlConnection(strCon))
            {
                SqlCommand sqlcomm = new SqlCommand(strComm, sqlcon);
                try
                {
                    sqlcon.Open();
                    result = sqlcomm.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MainMethod.InsertLog("ConnEWNAS.DeleteScrapWIPLog()-" + ex.Message);
                }
            }
            return result;
        }

        /// <summary>
        /// 取得防焊顯影生產日報表稽核結果
        /// </summary>
        /// <param name="StartDate">起始時間</param>
        /// <param name="EndDate">結束時間</param>
        /// <returns></returns>
        public DataTable ChkDevelopmentProductDailyReportLF(string StartDate, string EndDate)
        {
            var result = new DataTable();
            using (SqlConnection sqlcon = new SqlConnection(strCon))
            {
                SqlCommand sqlcomm = new SqlCommand(string.Empty, sqlcon);
                sqlcomm.CommandType = CommandType.StoredProcedure;
                sqlcomm.CommandText = "稽核防焊顯影生產日報表";
                sqlcomm.Parameters.Add("@FromDate", SqlDbType.DateTime);
                sqlcomm.Parameters.Add("@EndDate", SqlDbType.DateTime);
                sqlcomm.Parameters["@FromDate"].Value = StartDate;
                sqlcomm.Parameters["@EndDate"].Value = EndDate;
                try
                {
                    sqlcon.Open();
                    SqlDataReader reader = sqlcomm.ExecuteReader();
                    result.Load(reader);
                }
                catch (Exception ex)
                {
                    MainMethod.InsertLog("ChkDevelopmentProductDailyReportLF()-" + ex.Message);
                }
            }
            return result;
        }

        /// <summary>
        /// 取得乾膜顯影生產日報表稽核結果
        /// </summary>
        /// <param name="StartDate">起始時間</param>
        /// <param name="EndDate">結束時間</param>
        /// <returns></returns>
        public DataTable ChkDevelopmentProductDailyReportFF(string StartDate, string EndDate)
        {
            var result = new DataTable();
            using (SqlConnection sqlcon = new SqlConnection(strCon))
            {
                SqlCommand sqlcomm = new SqlCommand(string.Empty, sqlcon);
                sqlcomm.CommandType = CommandType.StoredProcedure;
                sqlcomm.CommandText = "稽核乾膜顯影生產日報表";
                sqlcomm.Parameters.Add("@FromDate", SqlDbType.DateTime);
                sqlcomm.Parameters.Add("@EndDate", SqlDbType.DateTime);
                sqlcomm.Parameters["@FromDate"].Value = StartDate;
                sqlcomm.Parameters["@EndDate"].Value = EndDate;
                try
                {
                    sqlcon.Open();
                    SqlDataReader reader = sqlcomm.ExecuteReader();
                    result.Load(reader);
                }
                catch (Exception ex)
                {
                    MainMethod.InsertLog("ChkDevelopmentProductDailyReportFF()-" + ex.Message);
                }
            }
            return result;
        }

        /// <summary>
        /// 取得乾膜壓膜生產日報表稽核結果
        /// </summary>
        /// <param name="StartDate">起始時間</param>
        /// <param name="EndDate">結束時間</param>
        /// <returns></returns>
        public DataTable Chk乾膜壓膜生產日報表(string StartDate, string EndDate)
        {
            var result = new DataTable();
            using (SqlConnection sqlcon = new SqlConnection(strCon))
            {
                SqlCommand sqlcomm = new SqlCommand(string.Empty, sqlcon);
                sqlcomm.CommandType = CommandType.StoredProcedure;
                sqlcomm.CommandText = "稽核乾膜壓膜生產日報表";
                sqlcomm.Parameters.Add("@FromDate", SqlDbType.DateTime);
                sqlcomm.Parameters.Add("@EndDate", SqlDbType.DateTime);
                sqlcomm.Parameters["@FromDate"].Value = StartDate;
                sqlcomm.Parameters["@EndDate"].Value = EndDate;
                try
                {
                    sqlcon.Open();
                    SqlDataReader reader = sqlcomm.ExecuteReader();
                    result.Load(reader);
                }
                catch (Exception ex)
                {
                    MainMethod.InsertLog("Chk乾膜壓膜生產日報表()-" + ex.Message);
                }
            }
            return result;
        }

        /// <summary>
        /// 取得乾膜AOI檢查日報表稽核結果
        /// </summary>
        /// <param name="StartDate">起始時間</param>
        /// <param name="EndDate">結束時間</param>
        /// <returns></returns>
        public DataTable Chk乾膜AOI檢查日報表(string StartDate, string EndDate)
        {
            var result = new DataTable();
            using (SqlConnection sqlcon = new SqlConnection(strCon))
            {
                SqlCommand sqlcomm = new SqlCommand(string.Empty, sqlcon);
                sqlcomm.CommandType = CommandType.StoredProcedure;
                sqlcomm.CommandText = "稽核乾膜AOI檢查日報表";
                sqlcomm.Parameters.Add("@FromDate", SqlDbType.DateTime);
                sqlcomm.Parameters.Add("@EndDate", SqlDbType.DateTime);
                sqlcomm.Parameters["@FromDate"].Value = StartDate;
                sqlcomm.Parameters["@EndDate"].Value = EndDate;
                try
                {
                    sqlcon.Open();
                    SqlDataReader reader = sqlcomm.ExecuteReader();
                    result.Load(reader);
                }
                catch (Exception ex)
                {
                    MainMethod.InsertLog("Chk乾膜AOI檢查日報表()-" + ex.Message);
                }
            }
            return result;
        }

        /// <summary>
        /// 取得成型V-CUT生產量測日報表稽核結果
        /// </summary>
        /// <param name="StartDate">起始時間</param>
        /// <param name="EndDate">結束時間</param>
        /// <returns></returns>
        public DataTable ChkVCutProductCheckRepay(string StartDate, string EndDate)
        {
            var result = new DataTable();
            using (SqlConnection sqlcon = new SqlConnection(strCon))
            {
                SqlCommand sqlcomm = new SqlCommand(string.Empty, sqlcon);
                sqlcomm.CommandType = CommandType.StoredProcedure;
                sqlcomm.CommandText = "稽核成型VCUT生產量測日報表";
                sqlcomm.Parameters.Add("@FromDate", SqlDbType.DateTime);
                sqlcomm.Parameters.Add("@EndDate", SqlDbType.DateTime);
                sqlcomm.Parameters["@FromDate"].Value = StartDate;
                sqlcomm.Parameters["@EndDate"].Value = EndDate;
                try
                {
                    sqlcon.Open();
                    SqlDataReader reader = sqlcomm.ExecuteReader();
                    result.Load(reader);
                }
                catch (Exception ex)
                {
                    MainMethod.InsertLog("ChkVCutProductCheckRepay()-" + ex.Message);
                }
            }
            return result;
        }

        /// <summary>
        /// 取得壓合PP裁切自主檢查日報表稽核結果
        /// </summary>
        /// <param name="StartDate">起始時間</param>
        /// <param name="EndDate">結束時間</param>
        /// <returns></returns>
        public DataTable ChkPPCutChkReport(string StartDate, string EndDate)
        {
            var result = new DataTable();
            using (SqlConnection sqlcon = new SqlConnection(strCon))
            {
                SqlCommand sqlcomm = new SqlCommand(string.Empty, sqlcon);
                sqlcomm.CommandType = CommandType.StoredProcedure;
                sqlcomm.CommandText = "稽核壓合PP裁切自主檢查日報表";
                sqlcomm.Parameters.Add("@FromDate", SqlDbType.DateTime);
                sqlcomm.Parameters.Add("@EndDate", SqlDbType.DateTime);
                sqlcomm.Parameters["@FromDate"].Value = StartDate;
                sqlcomm.Parameters["@EndDate"].Value = EndDate;
                try
                {
                    sqlcon.Open();
                    SqlDataReader reader = sqlcomm.ExecuteReader();
                    result.Load(reader);
                }
                catch (Exception ex)
                {
                    MainMethod.InsertLog("ChkPPCutChkReport()-" + ex.Message);
                }
            }
            return result;
        }
    }
}
