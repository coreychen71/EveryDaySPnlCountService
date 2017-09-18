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
        /// 取得當天已完成的製令單
        /// </summary>
        /// <returns></returns>
        public DataTable GetTodayFMEdIssue()
        {
            var Result = new DataTable();
            var strComm = "select t1.PaperNum '製令單號', CONVERT(char(19),t1.BuildDate,120) '作業時間', " +
                "t2.POTypeName '訂單種類', t1.LotNotes '批量種類', t1.PartNum '料號', t1.Revision '版序', " +
                "CONVERT(int, TotalPcs) '製令總PCS數' from FMEdIssueMain t1, FMEdPOType t2 " +
                "where t1.PaperDate = CONVERT(char(10), SYSDATETIME(), 120) and t1.Finished = 1 and " +
                "t1.POType = t2.POType order by t1.PaperNum";
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
    }
}