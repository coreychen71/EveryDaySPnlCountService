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
            using (StreamReader sr = new StreamReader(filepath, Encoding.Default))
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
                                //因應驗孔機的文字檔均使用空格當間隔，所以會遇到重複的Num名稱，需排除
                                if (aryStr[i] != "Num")
                                {
                                    //判斷欄位，將從文字檔讀出的 1-2孔、3-4孔、5-6孔、7-10孔、10孔以上等欄位乎略
                                    var q = 0;
                                    if (!int.TryParse(aryStr[i].ToString().Substring(0, 1), out q))
                                    {
                                        result.Columns.Add(aryStr[i]);
                                    }
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
                            /*
                             * 避免人員在驗孔機輸入料號時，料號後方有空一格再接續其它字串，導致在填值時，發生找不到資料行的情形，
                             * 因為會多出一欄，所以要檢查第二個欄位的值是否為0，文字檔裡的Betch Num欄位值均為0。
                            */
                            if (lstCol[1] == "0")
                            {
                                /*
                                 * 判斷i的數值是因為從文字檔抓出來的值，從第9~23都不要
                                 */
                                if (i < 9)
                                {
                                    //DataRow多加1是因為第一欄的操作者固定為null
                                    dr[i + 1] = lstCol[i];
                                }
                                else if(i > 23)
                                {
                                    //減掉14是因為 i 從 9~23 都不要，所以必需將 i 值減去略過的欄位，才能將正確的值填入DataTable的欄位
                                    dr[i - 14] = lstCol[i];
                                }
                            }
                        }
                    }
                    result.Rows.Add(dr);
                }
            }
            return result;
        }

    }
}
