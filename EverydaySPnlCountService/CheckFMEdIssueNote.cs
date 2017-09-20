using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EverydaySPnlCountService
{
    class CheckFMEdIssueNote
    {
        private string strCon = "server=ERP;database=EW;uid=jsis;pwd=JSIS";

        private string strConME = "server=EWNAS;database=ME;uid=me;pwd=2dae5na";

        /// <summary>
        /// 取得所有未交貨完畢的製令單
        /// </summary>
        /// <returns></returns>
        public DataTable GetAllFMEdIssue()
        {
            var Result = new DataTable();
            var strComm = "select t1.PaperNum '製令單號', t3.Item '項次', t3.PONum '備料單號', " +
                "CONVERT(char(19),t1.BuildDate,120) '作業時間', t2.POTypeName '訂單種類', t1.LotNotes '批量種類', " +
                "t1.PartNum '料號', t1.Revision '版序', CONVERT(int, TotalPcs) '製令總PCS數' " +
                "from FMEdIssueMain t1, FMEdPOType t2, FMEdIssuePO t3, SPOdOrderMain t4 " +
                "where t1.Finished = 1 and t1.POType = t2.POType and t1.PaperNum = t3.PaperNum and " +
                "t3.PONum = t4.PaperNum and t4.Finished = 1 order by t1.PaperNum, t3.Item";
            using (SqlConnection sqlcon = new SqlConnection(strCon))
            {
                using (SqlCommand sqlcomm = new SqlCommand(strComm, sqlcon))
                {
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

        /// <summary>
        /// 取得該料號工程工單途程的注意事項
        /// </summary>
        /// <param name="PartNum">料號</param>
        /// <param name="Revision">版序</param>
        /// <returns></returns>
        public DataTable GetIssueNote(string PartNum, string Revision)
        {
            var Result = new DataTable();
            var strComm = "select t3.LayerName '層別', REPLACE(t1.ProcCode,' ','') + ' ' + t2.ProcName '途程', " +
                "t1.Notes '注意事項' from EMOdLayerRoute t1, EMOdProcInfo t2, EMOdProdLayer t3 " +
                "where t1.PartNum = '" + PartNum + "' and t1.Revision = '" + Revision + "' and t1.Notes <> ' ' " +
                "and t1.ProcCode = t2.ProcCode and t1.PartNum = t3.PartNum and t1.Revision = t3.Revision " +
                "and t1.LayerId = t3.LayerId order by t1.LayerId, SerialNum asc";
            using (SqlConnection sqlcon = new SqlConnection(strCon))
            {
                using (SqlCommand sqlcomm = new SqlCommand(strComm, sqlcon))
                {
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

        /// <summary>
        /// 檢查製令單據是否已存在Log裡，若存在傳回true
        /// </summary>
        /// <param name="PaperNum">單據號碼</param>
        /// <returns></returns>
        public bool CheckIssueLog(string PaperNum)
        {
            var Result = false;
            var strComm = "select * from ChkFMEdIssueMainLog where PaperNum = '" + PaperNum + "'";
            using (SqlConnection sqlcon = new SqlConnection(strConME))
            {
                using (SqlCommand sqlcomm = new SqlCommand(strComm, sqlcon))
                {
                    try
                    {
                        sqlcon.Open();
                        SqlDataReader reader = sqlcomm.ExecuteReader();
                        if (reader.HasRows)
                        {
                            Result = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        MainMethod.InsertLog(ex.Message);
                    }
                }
            }
            return Result;
        }

        /// <summary>
        /// 寫入製令單據Log
        /// </summary>
        /// <param name="PaperNum">單據號碼</param>
        /// <returns></returns>
        public int InsertIssueLog(string PaperNum)
        {
            var Result = 0;
            var strComm = "insert into ChkFMEdIssueMainLog(PaperNum,LogTime) values('" + PaperNum + "','" +
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";
            using (SqlConnection sqlcon = new SqlConnection(strConME))
            {
                using (SqlCommand sqlcomm = new SqlCommand(strComm, sqlcon))
                {
                    try
                    {
                        sqlcon.Open();
                        Result = sqlcomm.ExecuteNonQuery();
                    }
                    catch(Exception ex)
                    {
                        MainMethod.InsertLog(ex.Message);
                    }
                }
            }
            return Result;
        }

        /// <summary>
        /// 取得Ewproject料號特殊追蹤事項清單
        /// </summary>
        /// <param name="PartNum">料號，只需要料號前7碼即可。</param>
        /// <returns></returns>
        public DataTable GetTeacePartNum(string PartNum)
        {
            var Result = new DataTable();
            var strComm = "select PARTNUM '料號', AFFAIR '特殊事項' from TRACEPARTNUM where PARTNUM like '" +
                PartNum + "%'";
            using (SqlConnection sqlcon = new SqlConnection(strConME))
            {
                using (SqlCommand sqlcomm = new SqlCommand(strComm, sqlcon))
                {
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