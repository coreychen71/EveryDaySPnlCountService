using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace EverydaySPnlCountService
{
    /// <summary>
    /// 操作 EWNAS CheckVCUT_JumpLog 資料表
    /// </summary>
    class CheckVCUT_JumpLog
    {
        private string strCon = "server=EWNAS;database=ME;uid=me;pwd=2dae5na";

        /// <summary>
        /// 檢查該筆紀錄是否已存在，存在傳回true
        /// </summary>
        /// <param name="LotNum"></param>
        /// <param name="ProcCode"></param>
        /// <param name="NextProcCode"></param>
        /// <returns></returns>
        public bool Check(string LotNum, string ProcCode, string NextProcCode)
        {
            var Result = false;
            var strComm = "select * from CheckVCUT_JumpLog where LotNum = '" + LotNum + "' and ProcCode='" + ProcCode +
                "' and NextProcCode='" + NextProcCode + "'";
            using (SqlConnection sqlcon = new SqlConnection(strCon))
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
        /// 寫入Log
        /// </summary>
        /// <param name="ProcCode">目前途程</param>
        /// <param name="LotNum">批號</param>
        /// <param name="PartNum">料號</param>
        /// <param name="Revision">版序</param>
        /// <param name="LayerId">層別</param>
        /// <param name="POP">型別</param>
        /// <param name="Qnty">數量</param>
        /// <param name="ReWork">重工</param>
        /// <param name="NextProcCode">下站途程</param>
        /// <returns></returns>
        public int Insert(string ProcCode, string LotNum, string PartNum, string Revision, string LayerId,
            string POP, string Qnty, string ReWork, string NextProcCode)
        {
            var Result = 0;
            var strComm = "insert into CheckVCUT_JumpLog(ProcCode,LotNum,PartNum,Revision,LayerId,POP,Qnty,ReWork," +
                "NextProcCode,LogTime) values('" + ProcCode + "','" + LotNum + "','" + PartNum + "','" + Revision + "','" +
                LayerId + "','" + POP + "','" + Qnty + "','" + ReWork + "','" + NextProcCode + "','" +
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";
            using (SqlConnection sqlcon = new SqlConnection(strCon))
            {
                using (SqlCommand sqlcomm = new SqlCommand(strComm, sqlcon))
                {
                    try
                    {
                        sqlcon.Open();
                        Result = sqlcomm.ExecuteNonQuery();
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
