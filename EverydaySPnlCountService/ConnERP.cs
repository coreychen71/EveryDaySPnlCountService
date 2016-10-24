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

        public ConnERP(string PaperDate)
        {
            pDate = PaperDate;
        }

        /// <summary>
        /// 取得特殊油墨製令單
        /// </summary>
        /// <param name="PaperDate">製令單日期</param>
        /// <returns></returns>
        public DataTable ChkPrintingInk()
        {
            var result = new DataTable();
            var strComm = "select A.PaperNum as '製令單號',A.PartNum as '料號',A.Revision as '版序',D.Length as '發料/L'," +
                "D.Width as '發料/W',B.POTypeName as '訂單種類',CONVERT(int, A.TotalPcs) as '製令總數量/Pcs'," +
                "CONVERT(int, A.SpareQnty) as '備品數/Pcs',CONVERT(char(10), A.ExpStkTime, 120) as '預計繳庫日'," +
                "A.LotNotes as '批量種類',C.LsmColor as '防焊TOP',C.LsmColorB as '防焊BOT',C.CharColor as '文字顏色 '" +
                "from FMEdIssueMain A,FMEdPOType B, EMOdProdInfo C,EMOdProdPOP D " +
                "where A.PaperDate = '" + pDate + "' and A.CancelDate is null and A.POType = B.POType and " +
                "A.PartNum = C.PartNum and A.Revision = C.Revision and A.PartNum = D.PartNum and " +
                "A.Revision = D.Revision and D.POP = 2 and " +
                "(SUBSTRING(A.PartNum, 9, 1) in ('H', 'I', 'J') or C.LsmColor like '[紫,白]%' or C.CharColor like '黑%') " +
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
                    MainMethod.InsertLog(ex.Message);
                }
            }
            return result;
        }
    }
}
