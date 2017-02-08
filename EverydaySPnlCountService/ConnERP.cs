using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace EverydaySPnlCountService
{
    class ConnERP
    {
        static string strCon = "server=ERP;database=EW;uid=jsis;pwd=JSIS";

        private string pDate = string.Empty;

        /// <summary>
        /// 指定單據日期
        /// </summary>
        /// <param name="PaperDate"></param>
        public ConnERP(string PaperDate)
        {
            pDate = PaperDate;
        }

        /// <summary>
        /// 取得特殊油墨製令單
        /// </summary>
        /// <returns></returns>
        public DataTable ChkPrintingInk()
        {
            var result = new DataTable();
            var strComm = "select A.PaperNum '製令單號', A.PartNum '料號', A.Revision '版序', D.Length '發料/L'," +
                "D.Width '發料/W',B.POTypeName '訂單種類', CONVERT(int, E.IOQnty) '製令PNL數'," +
                "CONVERT(char(10), A.ExpStkTime, 120) '預計繳庫日', A.LotNotes '批量種類', C.LsmColor '防焊TOP'," +
                "C.LsmColorB '防焊BOT', C.CharColor '文字顏色' from FMEdIssueMain A,FMEdPOType B, EMOdProdInfo C," +
                "EMOdProdPOP D, FMEdIssueSub E " +
                "where A.PaperDate = '" + pDate + "' and A.CancelDate is null and A.POType = B.POType and " +
                "A.PartNum = C.PartNum and A.Revision = C.Revision and A.PartNum = D.PartNum and " +
                "A.PaperNum = E.PaperNum and A.PartNum = E.PartNum and E.Item = 1 and " +
                "A.Revision = D.Revision and D.POP = 2 and " +
                "(SUBSTRING(A.PartNum, 9, 1) in ('H', 'I', 'J') or C.LsmColor like '[紅,紫,白,橘]%' or C.CharColor like '黑%') " +
                "order by A.BuildDate desc";
            using(SqlConnection sqlcon=new SqlConnection(strCon))
            {
                SqlCommand sqlcomm = new SqlCommand(strComm, sqlcon);
                try
                {
                    sqlcon.Open();
                    SqlDataReader read = sqlcomm.ExecuteReader();
                    result.Load(read);
                }
                catch (Exception ex)
                {
                    MainMethod.InsertLog("ConnERP.ChkPrintingInk()-" + ex.Message);
                }
            }
            return result;
        }

        /// <summary>
        /// 取得製令單據
        /// </summary>
        /// <returns></returns>
        public DataTable GetIssuePaper()
        {
            var result = new DataTable();
            var strComm = "select A.PaperNum '製令單號',CONVERT(char(10),A.PaperDate,120) '單據日期',A.PartNum '料號'," +
                "A.Revision '版序',B.POTypeName '訂單類型',CONVERT(int, A.TotalPcs) '製令總Pcs數'," +
                "CONVERT(char(10), A.ExpStkTime, 120) '期望繳庫日',A.FinishUser '審核人員' " +
                "from FMEdIssueMain A,FMEdPOType B where A.PaperDate = '" + pDate + "' and A.Finished = 1 and " +
                "A.POType = B.POType order by A.BuildDate,A.PartNum";
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
                    MainMethod.InsertLog("ConnERP.GetIssuePaper()-" + ex.Message);
                }
            }
            return result;
        }

        /// <summary>
        /// 取得現帳410&610預備報廢帳
        /// </summary>
        /// <returns></returns>
        public DataTable GetScrapWIP()
        {
            var result = new DataTable();
            var strComm = "sp_executesql N'exec FMEdProcSearch @P1,@P2',N'@P1 varchar(255),@P2 int'," +
                "'and (t1.ProcCode=''410'' or t1.ProcCode=''610'') and t1.LotStatus=2',0";
            using (SqlConnection sqlcon = new SqlConnection(strCon))
            {
                SqlCommand sqlcomm = new SqlCommand(strComm, sqlcon);
                try
                {
                    sqlcon.Open();
                    SqlDataReader reader = sqlcomm.ExecuteReader();
                    result.Load(reader);
                }
                catch(Exception ex)
                {
                    MainMethod.InsertLog("ConnERP.GetScrapWIP()-" + ex.Message);
                }
            }
            return result;
        }

        /// <summary>
        /// 檢查每日製令單的料號第9碼是否有含 I、V、Y (需使用樹脂)
        /// </summary>
        /// <returns></returns>
        public DataTable ChkIssueResin()
        {
            var result = new DataTable();
            var strComm = "select A.PaperNum '製令單號', A.PartNum '料號', A.Revision '版序', C.Length '發料/L'," +
                "C.Width '發料/W', B.POTypeName '訂單種類', CONVERT(int, D.IOQnty) '製令PNL數', CONVERT(char(10)," +
                "A.ExpStkTime, 120) '預計繳庫日', A.LotNotes '批量種類' " +
                "from FMEdIssueMain A,FMEdPOType B, EMOdProdPOP C,FMEdIssueSub D " +
                "where A.PaperDate = '" + pDate + "' and A.CancelDate is null and A.POType = B.POType and " +
                "A.PartNum = C.PartNum and A.Revision = C.Revision and C.POP = 2 and " +
                "A.PaperNum = D.PaperNum and A.PartNum = D.PartNum and D.Item = 1 and " +
                "SUBSTRING(A.PartNum, 9, 1) in ('I', 'V', 'Y') " +
                "order by A.PartNum";
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
                    MainMethod.InsertLog("C.ERP.ChkIssueResin()-" + ex.Message);
                }
            }
            return result;
        }

        /// <summary>
        /// 取得批號資訊明細
        /// </summary>
        /// <param name="LotNum">批號</param>
        /// <returns></returns>
        public static DataTable GetLotInfoSreach(string LotNum)
        {
            var result = new DataTable();
            var strComm = "sp_executesql N'exec FMEdLotInfoSearch @P1',N'@P1 varchar(255)'," +
                "'and t1.LotNum=''" + LotNum + "'''";
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
                    MainMethod.InsertLog("ConnEWNAS.GetLotInfoSreach()-" + ex.Message);
                }
            }
            return result;
        }

        /// <summary>
        /// 檢查製令單號是否有在傑偲調帳子表裡，若有表示有增帳。
        /// 有就回傳true
        /// </summary>
        /// <param name="MotherIssueLotNum">母製令單號</param>
        /// <returns>bool</returns>
        public static bool ChkFMEdTuneSub(string MotherIssueLotNum)
        {
            var result = false;
            var strComm = "select * from FMEdTuneSub where MotherIssueNum='" + MotherIssueLotNum + "'";
            using (SqlConnection sqlcon = new SqlConnection(strCon))
            {
                SqlCommand sqlcomm = new SqlCommand(strComm, sqlcon);
                try
                {
                    sqlcon.Open();
                    SqlDataReader reader = sqlcomm.ExecuteReader();
                    if (reader.HasRows)
                    {
                        result = true;
                    }
                }
                catch (Exception ex)
                {
                    MainMethod.InsertLog("ConnEWNAS.ChkFMEdTuneSub()-" + ex.Message);
                }
            }
            return result;
        }

        /// <summary>
        /// 檢查是否有此批號的批號報廢單(狀態=>報廢轉正常)
        /// 生管人員幫現場預報帳增帳的方式
        /// </summary>
        /// <param name="LotNum">批號</param>
        /// <returns></returns>
        public static bool ChkFMEdStatusScrap(string LotNum)
        {
            var result = false;
            var strComm = "select A.PaperNum,A.PaperDate,B.LotNum,B.PartNum,B.Revision,B.LayerId,B.POP,B.Qnty," +
                "B.IsScrap,C.FinishedName,A.FinishUser from FMEdStatusScrapMain A, FMEdStatusScrapSub B," +
                "CURdPaperFinished C where B.LotNum = '" + LotNum + "' and B.IsScrap = 0 and A.Finished = 1 and " +
                "B.PaperNum = A.PaperNum and A.Finished = C.Finished";
            using (SqlConnection sqlcon = new SqlConnection(strCon))
            {
                SqlCommand sqlcomm = new SqlCommand(strComm, sqlcon);
                try
                {
                    sqlcon.Open();
                    SqlDataReader Reader = sqlcomm.ExecuteReader();
                    if (Reader.HasRows)
                    {
                        result = true;
                    }
                }
                catch (Exception ex)
                {
                    MainMethod.InsertLog("ConnEWNAS.ChkFMEdTuneSub()-" + ex.Message);
                }
            }
            return result;
        }
    }
}
