using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EverydaySPnlCountService
{
    class TXTtoTable
    {
        string filepath = string.Empty;

        public TXTtoTable(string FilePath)
        {
            filepath = FilePath;
        }

        public DataTable GetTable()
        {
            var result = new DataTable();
            using (StreamReader sr = new StreamReader(filepath))
            {
                var str = string.Empty;
                while ((str = sr.ReadLine()) != null)
                {
                    var dr = result.NewRow();
                    string[] aryStr = str.Split(' ');
                    //若DataTable沒有任何資料行，則先將文字檔的第一行填入當做DataTable第一行的欄位名稱
                    if (result.Rows.Count == 0)
                    {
                        for (int i = 0; i < aryStr.Length; i++)
                        {
                            if (aryStr[i] != "")
                            {
                                //因應鑽孔驗孔機的文字紀錄檔均使用空格當間隔，所以會遇到重複的Num名稱，需排除
                                if (aryStr[i] != "Num")
                                {
                                    result.Columns.Add(aryStr[i]);
                                }
                            }
                        }
                    }
                    //反之，若已有欄位名稱資料行，就開始填入資料
                    else
                    {
                        //先建一個List
                        var lstCol = new List<string>();

                        //將陣列中有值的項次填入List中
                        foreach (string strCol in aryStr)
                        {
                            if (strCol != "")
                            {
                                lstCol.Add(strCol);
                            }
                        }

                        //再將List的值填入DataRow
                        for(int i = 0; i < lstCol.Count; i++)
                        {
                            //DataRow多加1是因為第一欄的User固定為null
                            dr[i + 1] = lstCol[i];
                        }
                    }
                    result.Rows.Add(dr);
                }
            }
            return result;
        }

    }
}
