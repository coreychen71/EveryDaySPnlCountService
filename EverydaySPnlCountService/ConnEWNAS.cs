using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace EverydaySPnlCountService
{
    class ConnEWNAS
    {
        static string strCon = "server=EWNAS;database=ME;uid=me;pwd=2dae5na";

        /// <summary>
        /// 取得Ewproject每日未結案客訴事項
        /// </summary>
        /// <returns></returns>
        public DataTable EveryDayCustomerComplaint()
        {
            var result = new DataTable();
            var strComm = "select custermerid '客戶別',CONVERT(char(10),cpldate,120) '客訴日期',datecode '週期'," +
                "PARTNUM '料號',custermerpartnum '客戶料號',QNTY '數量',defectname '原因',discribe '描述'," +
                "dutydp '責任單位',dutyman '責任人',CONVERT(char(10), requirdate, 120) '預計完成日',writeman '登記人員'," +
                "CONVERT(char(10), writedate, 120) '登記日期' from qacustermercpl where stateflag is null " +
                "order by cpldate,custermerid,PARTNUM";
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
    }
}
