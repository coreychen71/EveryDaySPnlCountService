﻿using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Data;
using System.Text;
using System.Timers;
using EWPCB_SPnlCountClass;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;

namespace EverydaySPnlCountService
{
    class MainMethod
    {
        //依設定時間執行用Timer
        private Timer timer = new Timer();
        //固定每5分鐘就執行的Timer
        private Timer timer2 = new Timer();
        //設定Timer1間隔時間為1分鐘
        private double timerInterval = 60 * 1000;
        //設定Timer2間隔時間為5分鐘
        private double timerInterval2 = 5 * 60 * 1000;
        //private DateTime setTime;
        private DateTime nowTime;
        private string datetimeFormat = "yyyy-MM-dd HH:mm:ss";
        private string SaveFile="";

        public MainMethod()
        {
            timer.Interval = timerInterval;
            timer.AutoReset = true;
            timer.Elapsed += Timer_Elapsed;
            timer2.Interval = timerInterval2;
            timer2.AutoReset = true;
            timer2.Elapsed += Timer2_Elapsed;
            timer2.Start();
            //SPnlCountRun();
            //GetEveryDayCustomerComplaint();
            //ChkProductDailyReport();
            //ChkDrillHole();
            //ChkPrintingInk();
            //ChkIssueAndScrapWIP();
            //ChkScrapWIPLog();
            //ChkIQC每日量測申報紀錄();
            //ChkDepotStock();
            //ChkVCUT_Jump();
            //ChkFMEdIssueNote();
            //ChkTracePartNum();
            //ChkDrill2ndProcByMonth();
            //CheckAcknowledgmentIn650();
            //CheckAcknowledgmentAppointDate();
            //HPSdGetSpecialBreakDay();
        }

        #region Timer Block
        /// <summary>
        /// Timer Start
        /// </summary>
        public void Start()
        {
            timer.Start();
        }

        /// <summary>
        /// Timer Stop
        /// </summary>
        public void Stop()
        {
            timer.Stop();
        }

        /// <summary>
        /// Timer 觸發事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            //每日06:30進行SPnl數統計(TXT)
            if (CheckTime("06:30:00", "06:30:59"))
            {
                Stop();
                SPnlCountRun();
                Start();
            }
            //每日07:00進行品保客訴待處理(未逾期)清單(Excel)、IQC每日量測申報稽核
            else if (CheckTime("07:00:00", "07:00:59"))
            {
                Stop();
                GetEveryDayCustomerComplaint();
                ChkIQC每日量測申報紀錄();
                Start();
            }
            //每日08:00查詢倉庫不足安全庫存量物料
            else if (CheckTime("08:00:00", "08:00:59"))
            {
                Stop();
                ChkDepotStock();
                Start();
            }
            //每日08:20進行輔助系統生產日報表稽核
            else if (CheckTime("08:20:00", "08:20:59"))
            {
                Stop();
                ChkProductDailyReport();
                Start();
            }
            //每日09:00進行二次鑽料號工單漏420途程通知
            else if (CheckTime("09:00:00", "09:00:59"))
            {
                Stop();
                ChkDrill2ndProcByMonth();
                Start();
            }
            //每日11:30進行未交貨完畢製令單特殊事項檢查
            else if (CheckTime("11:30:00", "11:30:59"))
            {
                Stop();
                ChkTracePartNum();
                Start();
            }
            //每日16:00進行待修設備清單(Excel)
            else if (CheckTime("16:00:00", "16:00:59"))
            {
                Stop();
                EquipMaintainRun();
                Start();
            }
            //每日16:30查詢倉庫不足安全庫存量物料、未交貨完畢製令單特殊事項檢查、二次鑽料號工單漏420途程通知
            else if (CheckTime("16:30:00", "16:30:59"))
            {
                Stop();
                ChkDepotStock();
                ChkTracePartNum();
                ChkDrill2ndProcByMonth();
                Start();
            }
            //每日22:30進行未交貨完畢製令單特殊事項檢查
            else if (CheckTime("22:30:00", "22:30:59"))
            {
                Stop();
                ChkTracePartNum();
                Start();
            }
            //每日00:10進行每日製令稽核現帳預報廢清單(Excel)、特休假即將到期人員查詢(Exce)-每月1號才檢查
            else if (CheckTime("00:10:00", "00:10:59"))
            {
                Stop();
                if (DateTime.Now.Day == 1)
                {
                    HPSdGetSpecialBreakDay();
                }
                ChkIssueAndScrapWIP();
                Start();
            }
            //每日00:15進行每日製令稽核現帳預報廢未增帳清單(Excel)
            else if (CheckTime("00:15:00", "00:15:59"))
            {
                Stop();
                ChkScrapWIPLog();
                Start();
            }
            //每日00:20進行鑽孔申報比對驗孔紀錄(TXT)
            else if (CheckTime("00:20:00", "00:20:59"))
            {
                Stop();
                ChkDrillHole();
                Start();
            }
            //每日00: 30進行製令單料號檢查是否有特殊油墨與樹脂需求(TXT)
            else if (CheckTime("00:30:00", "00:30:59"))
            {
                Stop();
                ChkPrintingInk();
                Start();
            }
        }

        /// <summary>
        /// Timer2 觸發事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Timer2_Elapsed(object sender, ElapsedEventArgs e)
        {
            timer2.Stop();
            ChkVCUT_Jump();
            CheckAcknowledgmentIn650();
            CheckAcknowledgmentAppointDate();
            //2017/09/19 停用，主管反應信件過多
            //ChkFMEdIssueNote();
            timer2.Start();
        }
        #endregion

        /// <summary>
        /// 寫入Log
        /// </summary>
        /// <param name="Msg">要寫入的訊息</param>
        public static void InsertLog(string Msg)
        {
            if (!Directory.Exists(@"C:\EveryDaySPnlCountService"))
            {
                DirectoryInfo dir = new DirectoryInfo(@"C:\EveryDaySPnlCountService");
                dir.Create();
            }
            string LogPath = @"C:\EveryDaySPnlCountService\EveryDaySPnlCountLog.txt";
            StreamWriter writerLog;
            //此方式為每次寫入時，持續寫入，不會覆蓋原本內容
            writerLog = File.AppendText(LogPath);
            writerLog.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + Msg + "\r\n");
            writerLog.Flush();
            writerLog.Close();
            writerLog.Dispose();
        }

        /// <summary>
        /// 檢查目前的系統時間是否符合輸入的二者時間區間內
        /// </summary>
        /// <param name="time1">第一時間</param>
        /// <param name="time2">第二時間</param>
        /// <returns></returns>
        private bool CheckTime(string time1,string time2)
        {
            //##### 2016/07/30 改用判斷是否在指定的時間區段內 #####//
            bool result = false;
            nowTime = DateTime.Now;
            DateTime t1 = DateTime.Parse(time1);
            DateTime t2 = DateTime.Parse(time2);
            if (nowTime >= t1 && nowTime <= t2)
            {
                result = true;
            }
            //else
            //{
            //    //### 初期查看時間判斷誤差值使用 ###//
            //    writerLog.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "   CheckTime false_" + nowTime);
            //    writerLog.Flush();
            //}
            return result;
            #region 原先一開始所使用的秒數差距判斷(已停用)
            //##### 原先一開始所使用的秒數差距判斷 #####//
            /*
            setTime = DateTime.Parse(settime);
            double interval = nowTime.Subtract(setTime).TotalSeconds;
            //差距在60秒內均為true
            if (interval >= -60 && interval <= 60)
            {
                result = true;
                writerLog.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "   CheckTime true_" + nowTime +
                    "_interval=" + interval);
                writerLog.Flush();
            }
            */
            #endregion
        }

        /// <summary>
        /// 取得一個月前至當天日期的每日預估出貨SPnl數量
        /// </summary>
        private void SPnlCountRun()
        {
            try
            {
                StreamWriter writerResult;
                SaveFile = Path.GetTempPath() + "SPnlCount.txt";
                //### 2016/07/21 依總經理指示，再將起始日往前推1個月。 ###//
                DateTime startDate = DateTime.Now.AddMonths(-1);
                SPnlCount sc = new SPnlCount();
                using (DataTable result = sc.SPnlCountRun(startDate))
                {
                    //此方式為每次寫入時，均會直接覆蓋掉原本內容
                    writerResult = new StreamWriter(SaveFile);
                    writerResult.WriteLine("日期        SPnl數量" + "\r\n");
                    for (int i = 0; i < result.Rows.Count; i++)
                    {
                        writerResult.WriteLine(result.Rows[i][0].ToString() + "  " +
                            result.Rows[i][1].ToString() + "\r\n");
                    }
                    writerResult.Flush();
                    writerResult.Close();
                    writerResult.Dispose();
                }
            }
            catch (Exception ex)
            {
                InsertLog(DateTime.Now.ToString(datetimeFormat) + "  SPnlCountRun()" + ex.Message + "\r\n");
            }
            SendMail("sm4@ewpcb.com.tw", "每日預估入庫SPnl數量統計", "spnlcount@ewpcb.com.tw",
                    DateTime.Now.ToString("yyyy-MM-dd") + " 每日預估入庫SPnl數量統計！",
                    "每日預估入庫SPnl數量統計，請詳閱附件。" + "<br/>" + "<br/>" +
                    "-----此封郵件由系統所寄出，請勿直接回覆！-----", SaveFile);
        }

        /// <summary>
        /// 取得每日設備待維修案件
        /// </summary>
        private void EquipMaintainRun()
        {
            try
            {
                SaveFile = Path.GetTempPath() + "EquipMaintain.xls";
                DataTable[] result = new DataTable[] { EquipMaintain.GetWaitMaintain() };
                string[] strSheet = new string[] { "每日待修設備清單" };
                DataTableToExcel(result, strSheet, SaveFile);
            }
            catch (Exception ex)
            {
                InsertLog(DateTime.Now.ToString(datetimeFormat) + "  EquipMaintainRun()" + ex.Message + "\r\n");
            }
            SendMail("sm4@ewpcb.com.tw", "每日待修設備清單", "equip@ewpcb.com.tw",
                    DateTime.Now.ToString("yyyy-MM-dd") + " 每日待修設備清單！",
                    "每日待修設備清單，請詳閱附件。" + "<br/>" + "<br/>" +
                    "-----此封郵件由系統所寄出，請勿直接回覆！-----", SaveFile);
        }

        /// <summary>
        /// 稽核鑽孔課驗孔機每日驗孔紀錄
        /// </summary>
        private void ChkDrillHole()
        {
            //設定要扣除的天數
            var decreaseDate = -1;
            //var srcPath = @"\\192.168.1.200\DailyReport4\";
            var srcPath = @"E:\DailyReport4\";
            try
            {
                StreamWriter writerResult;
                SaveFile = Path.GetTempPath() + DateTime.Now.AddDays(decreaseDate).ToString("yyyy-MM-dd") +
                    "驗孔數量稽核清單.txt";
                writerResult = new StreamWriter(SaveFile);
                writerResult.WriteLine("批號\t料號\t數量\t驗板數量\t驗板時間(起)\t驗板時間(迄)");
                DFCheckHoleRecord chk = new DFCheckHoleRecord();
                var ewTB = chk.GetDFRecord(DateTime.Now.AddDays(decreaseDate).ToString("yyyy-MM-dd"));
                TXTtoTable loadingTxt = new TXTtoTable(srcPath +
                DateTime.Now.AddDays(decreaseDate).ToString("yyyyMMdd") + ".txt");
                var chkTB = loadingTxt.GetTable();

                #region 開始稽核數量
                foreach (DataRow row in ewTB.Rows)
                {
                    var chkCount = 0;
                    var StartTime = string.Empty;
                    var EndTime = string.Empty;
                    foreach (DataRow chkrow in chkTB.Rows)
                    {
                        try
                        {
                            //chkrow[1]=料號、chkrow[3]=驗板開始時間、chkrow[5]=檢驗板數、chkrow[13]=驗板結束時間
                            if (chkrow[1].ToString() != "")
                            {
                                //檢查料號是否有輸入不完全的
                                if (chkrow[1].ToString().Length >= 11)
                                {
                                    if (row["料號"].ToString().Trim() ==
                                        chkrow[1].ToString().ToUpper().Substring(0, 11))
                                    {
                                        chkCount += Convert.ToInt32(chkrow[5]);
                                        StartTime = chkrow[3].ToString();
                                        EndTime = chkrow[13].ToString();
                                    }
                                }
                                else
                                {
                                    if (row["料號"].ToString().Trim().Substring(0, 8) ==
                                        chkrow[1].ToString().ToUpper().Substring(0, 8))
                                    {
                                        chkCount += Convert.ToInt32(chkrow[5]);
                                        StartTime = chkrow[3].ToString();
                                        EndTime = chkrow[13].ToString();
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            InsertLog(DateTime.Now.ToString(datetimeFormat) + "  " + ex.Message + "\r" +
                                row["料號"].ToString().Trim() + "+" + chkrow[1].ToString() + "\r");
                        }
                    }
                    if (Convert.ToInt32(row["數量"]) > chkCount)
                    {
                        writerResult.WriteLine(row["批號"].ToString() + "\t" + row["料號"].ToString() + "\t" +
                            row["數量"].ToString() + "\t" + Convert.ToString(chkCount) + "\t" + StartTime + "\t" +
                            EndTime);
                    }
                }
                #endregion

                #region 105/10/27 停用識別是否Ewproject有申報紀錄但卻未驗板的功能
                //writerResult.WriteLine();
                //writerResult.WriteLine();
                //writerResult.WriteLine("========== 有申報Ewproject，但未有驗孔紀錄的料號 ==========");
                //writerResult.WriteLine();
                //writerResult.WriteLine("批號\t料號\t數量\t開始時間\t結束時間\t人員");
                //#region 檢查是否有在Ewproject申報的批號，卻沒有進行驗板
                //foreach (DataRow sRow in ewTB.Rows)
                //{
                //    //若不符合，就把result+1，等內層迴圈跑完，result等於chkTB的筆數，就表示該筆料號未在驗孔機LOG出現過
                //    var result = 0;
                //    var chkTBrow = chkTB.Rows.Count;
                //    foreach (DataRow cRow in chkTB.Rows)
                //    {
                //        if (cRow[1].ToString() != "")
                //        {
                //            if (cRow[1].ToString().Length >= 11)
                //            {
                //                if (cRow[1].ToString().ToUpper().Substring(0, 11).Contains(sRow["料號"].ToString().Trim()))
                //                {
                //                    break;
                //                }
                //                else
                //                {
                //                    result++;
                //                }
                //            }
                //            else
                //            {
                //                if (cRow[1].ToString().ToUpper().Substring(0, 8).Contains(sRow["料號"].ToString().Trim()))
                //                {
                //                    break;
                //                }
                //                else
                //                {
                //                    result++;
                //                }
                //            }
                //        }
                //    }
                //    if (chkTBrow == result)
                //    {
                //        writerResult.WriteLine(sRow["批號"].ToString().Trim() + "\t" +
                //            sRow["料號"].ToString().Trim() + "\t" +
                //            sRow["數量"].ToString().Trim() + "\t" +
                //            sRow["開始時間"].ToString().Trim() + "\t" +
                //            sRow["結束時間"].ToString().Trim() + "\t" +
                //            sRow["人員"].ToString().Trim());
                //    }
                //}
                #endregion

                writerResult.Flush();
                writerResult.Close();
                writerResult.Dispose();
            }
            catch (Exception ex)
            {
                InsertLog(DateTime.Now.ToString(datetimeFormat) + "  ChkDrillHole()" + ex.Message + "\n\r");
            }
            SendMail("sm4@ewpcb.com.tw", "鑽孔每日驗孔數量稽核清單", "checkhole@ewpcb.com.tw",
                    DateTime.Now.AddDays(decreaseDate).ToString("yyyy-MM-dd") + " 驗孔數量稽核清單！",
                    "鑽孔每日驗孔數量稽核清單，請詳閱附件。" + "<br/>" + "<br/>" +
                    "-----此封郵件由系統所寄出，請勿直接回覆！-----", SaveFile);
        }

        /// <summary>
        /// 檢查每日製令料號的防焊油墨是否有特殊油墨需求
        /// </summary>
        private void ChkPrintingInk()
        {
            StreamWriter writerResult;
            SaveFile = Path.GetTempPath() + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + "製令單特殊油墨.txt";
            writerResult = new StreamWriter(SaveFile);
            var result = new DataTable();
            ConnERP ce = new ConnERP();
            ce.PaperDate = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
            result = ce.ChkPrintingInk();
            writerResult.WriteLine(
                string.Format("{0}\t{1}\t\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}\t\t{10}\t\t{11}",
                result.Columns[0].ColumnName,
                result.Columns[1].ColumnName,
                result.Columns[2].ColumnName,
                result.Columns[3].ColumnName,
                result.Columns[4].ColumnName,
                result.Columns[5].ColumnName,
                result.Columns[6].ColumnName,
                result.Columns[7].ColumnName,
                result.Columns[8].ColumnName,
                result.Columns[9].ColumnName,
                result.Columns[10].ColumnName,
                result.Columns[11].ColumnName));
            foreach (DataRow row in result.Rows)
            {
                writerResult.WriteLine(
                    string.Format("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t\t\t{6}\t{7}\t{8}\t\t{9}\t{10}\t{11}",
                    row[0].ToString().Trim(),
                    row[1].ToString().Trim(),
                    row[2].ToString().Trim(),
                    row[3].ToString().Trim(),
                    row[4].ToString().Trim(),
                    row[5].ToString().Trim(),
                    row[6].ToString().Trim(),
                    row[7].ToString().Trim(),
                    row[8].ToString().Trim(),
                    row[9].ToString().Trim(),
                    row[10].ToString().Trim(),
                    row[11].ToString().Trim()));
            }

            //2016-12-07 工程課長要求，再加入料號需使用到樹脂的檢查
            result.Clear();
            result = ce.ChkIssueResin();
            writerResult.WriteLine();
            writerResult.WriteLine();
            writerResult.WriteLine("==================== 樹脂清單 ====================");
            writerResult.WriteLine();
            writerResult.WriteLine(string.Format("{0}\t{1}\t\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}\t{8}",
                result.Columns[0].ColumnName,
                result.Columns[1].ColumnName,
                result.Columns[2].ColumnName,
                result.Columns[3].ColumnName,
                result.Columns[4].ColumnName,
                result.Columns[5].ColumnName,
                result.Columns[6].ColumnName,
                result.Columns[7].ColumnName,
                result.Columns[8].ColumnName));
            foreach (DataRow row in result.Rows)
            {
                writerResult.WriteLine(string.Format("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t\t\t{6}\t{7}\t{8}",
                    row[0].ToString().Trim(),
                    row[1].ToString().Trim(),
                    row[2].ToString().Trim(),
                    row[3].ToString().Trim(),
                    row[4].ToString().Trim(),
                    row[5].ToString().Trim(),
                    row[6].ToString().Trim(),
                    row[7].ToString().Trim(),
                    row[8].ToString().Trim()));
            }
            writerResult.Flush();
            writerResult.Close();
            writerResult.Dispose();
            result.Dispose();
            SendMail("sm4@ewpcb.com.tw", "製令單特殊油墨清單", "chkprintingink@ewpcb.com.tw",
                DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + " 製令單特殊油墨清單！",
                "製令單特殊油墨清單，請詳閱附件。" + "<br/>" + "<br/>" +
                "-----此封郵件由系統所寄出，請勿直接回覆！-----", SaveFile);
        }

        /// <summary>
        /// 取得每日待處理客訴事件(未逾期)
        /// </summary>
        private void GetEveryDayCustomerComplaint()
        {
            try
            {
                SaveFile = Path.GetTempPath() + DateTime.Now.ToString("yyyy-MM-dd") + "未逾期.xls";
                ConnEWNAS ce = new ConnEWNAS();
                DataTable[] result = new DataTable[] { ce.EveryDayCustomerComplaint() };
                string[] strSheet = new string[] { "品保待處理客訴通知(未逾期)" };
                DataTableToExcel(result, strSheet, SaveFile);
                #region 輸出文字檔(已停用)
                //輸出文字檔
                //writerResult = new StreamWriter(SaveFile);
                //writerResult.WriteLine("客戶別\t客訴日期\t週期\t料號\t客戶料號\t數量\t原因\t描述\t責任單位\t預計完成日\t" +
                //    "登記人員\t登記日期");
                //ConnEWNAS ce = new ConnEWNAS();
                //var srcTB = ce.EveryDayCustomerComplaint();
                //foreach (DataRow row in srcTB.Rows)
                //{
                //    writerResult.WriteLine(string.Format(
                //        "{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}\t{10}\t{11}",
                //        row["客戶別"].ToString().Trim(),
                //        row["客訴日期"].ToString().Trim(),
                //        row["週期"].ToString().Trim(),
                //        row["料號"].ToString().Trim(),
                //        row["客戶料號"].ToString().Trim(),
                //        row["數量"].ToString().Trim(),
                //        row["原因"].ToString().Trim(),
                //        row["描述"].ToString().Trim(),
                //        row["責任單位"].ToString().Trim(),
                //        row["預計完成日"].ToString().Trim(),
                //        row["登記人員"].ToString().Trim(),
                //        row["登記日期"].ToString().Trim()));
                //}
                //writerResult.Flush();
                //writerResult.Close();
                #endregion
            }
            catch (Exception ex)
            {
                InsertLog(DateTime.Now.ToString(datetimeFormat) + "  GetEveryDayCustomerComplaint()" + 
                    ex.Message + "\n\r");
            }
            SendMail("sm4@ewpcb.com.tw", "品保待處理客訴通知(未逾期)", "qagroup@ewpcb.com.tw",
                    DateTime.Now.ToString("yyyy-MM-dd") + " 品保待處理客訴通知(未逾期)",
                    "品保待處理客訴通知(未逾期)，請詳閱附件。" + "<br/>" + "<br/>" +
                    "-----此封郵件由系統所寄出，請勿直接回覆！-----", SaveFile);
        }

        /// <summary>
        /// 檢查防焊預報廢現帳清單是否有被下新製令單
        /// </summary>
        private void ChkIssueAndScrapWIP()
        {
            var result = new DataTable();
            var issue = new DataTable();
            var scrap = new DataTable();
            result.Columns.Add("製令單號");
            result.Columns.Add("單據日期");
            result.Columns.Add("料號");
            result.Columns.Add("版序");
            result.Columns.Add("訂單類型");
            result.Columns.Add("製令總Pcs數");
            result.Columns.Add("期望繳庫日");
            result.Columns.Add("審核人員");
            result.Columns.Add("提出製程");
            result.Columns.Add("代碼");
            result.Columns.Add("批號");
            result.Columns.Add("狀態");
            result.Columns.Add("階段名稱");
            result.Columns.Add("型狀");
            result.Columns.Add("排版");
            result.Columns.Add("預報廢數量");
            result.Columns.Add("批量種類");
            try
            {
                ConnERP ce = new ConnERP();
                ce.PaperDate = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
                issue = ce.GetIssuePaper();
                scrap = ce.GetScrapWIP();
                foreach (DataRow issueRow in issue.Rows)
                {
                    foreach (DataRow scrapRow in scrap.Rows)
                    {
                        if (issueRow["料號"].ToString().Trim() == scrapRow["PartNum"].ToString().Trim() &
                            issueRow["版序"].ToString().Trim() == scrapRow["Revision"].ToString().Trim())
                        {
                            DataRow row = result.NewRow();
                            row["製令單號"] = issueRow["製令單號"].ToString().Trim();
                            row["單據日期"] = issueRow["單據日期"].ToString().Trim();
                            row["料號"] = issueRow["料號"].ToString().Trim();
                            row["版序"] = issueRow["版序"].ToString().Trim();
                            row["訂單類型"] = issueRow["訂單類型"].ToString().Trim();
                            row["製令總Pcs數"] = issueRow["製令總Pcs數"].ToString().Trim();
                            row["期望繳庫日"] = issueRow["期望繳庫日"].ToString().Trim();
                            row["審核人員"] = issueRow["審核人員"].ToString().Trim();
                            row["提出製程"] = scrapRow["ProcName"].ToString().Trim();
                            row["代碼"] = scrapRow["ProcCode"].ToString().Trim();
                            row["批號"] = scrapRow["LotNum"].ToString().Trim();
                            row["狀態"] = scrapRow["LotStatusName"].ToString().Trim();
                            row["階段名稱"] = scrapRow["LayerName"].ToString().Trim();
                            row["型狀"] = scrapRow["POPName"].ToString().Trim();
                            row["排版"] = scrapRow["POP_SP"].ToString().Trim();
                            row["預報廢數量"] = Convert.ToInt32(scrapRow["Qnty"]);
                            row["批量種類"] = scrapRow["LotNotes"].ToString().Trim();
                            result.Rows.Add(row);
                        }
                    }
                }
                ConnEWNAS cewnas = new ConnEWNAS();
                foreach (DataRow row in result.Rows)
                {
                    cewnas.InsertScrapWIPLog(
                        row["製令單號"].ToString().Trim(),
                        row["單據日期"].ToString().Trim(),
                        row["料號"].ToString().Trim(),
                        row["版序"].ToString().Trim(),
                        row["訂單類型"].ToString().Trim(),
                        row["製令總Pcs數"].ToString().Trim(),
                        row["期望繳庫日"].ToString().Trim(),
                        row["審核人員"].ToString().Trim(),
                        row["提出製程"].ToString().Trim(),
                        row["代碼"].ToString().Trim(),
                        row["批號"].ToString().Trim(),
                        row["狀態"].ToString().Trim(),
                        row["階段名稱"].ToString().Trim(),
                        row["型狀"].ToString().Trim(),
                        row["排版"].ToString().Trim(),
                        row["預報廢數量"].ToString().Trim(),
                        row["批量種類"].ToString().Trim());
                }
                SaveFile = Path.GetTempPath() + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") +
                    "清單.xls";
                DataTable[] resultTB = new DataTable[] { result };
                string[] strSheet = new string[] { "製令稽核防焊現帳預報廢清單" };
                DataTableToExcel(resultTB, strSheet, SaveFile);
                issue.Dispose();
                scrap.Dispose();
            }
            catch (Exception ex)
            {
                InsertLog(DateTime.Now.ToString(datetimeFormat) + "  ChkIssueAndScrapWIP()" + 
                    ex.Message + "\n\r");
            }
            SendMail("sm4@ewpcb.com.tw", "每日製令稽核防焊現帳預報廢清單", "chkissuescrapwip@ewpcb.com.tw",
                    DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + " 製令稽核防焊現帳預報廢清單",
                    "製令稽核防焊現帳預報廢清單，請詳閱附件。" + "<br/>" + "<br/>" +
                    "-----此封郵件由系統所寄出，請勿直接回覆！-----", SaveFile);
        }

        /// <summary>
        /// 檢查防焊預報廢現帳LOG裡的製令單日期是否已達到二天
        /// 2天內未增帳就發信通知，已增帳就從LOG裡移除
        /// </summary>
        private void ChkScrapWIPLog()
        {
            ConnEWNAS ewnas = new ConnEWNAS();
            var ScrapWIPLog = ewnas.GetScrapWIPLog();
            var strNowDate = DateTime.Now;
            var result = ScrapWIPLog.Clone();
            foreach (DataRow row in ScrapWIPLog.Rows)
            {
                var IssueDate = DateTime.Parse(row["IssuePaperDate"].ToString().Trim());
                ConnERP ce = new ConnERP();
                if (IssueDate.AddDays(2) <= strNowDate)
                {
                    //先檢查是否已增帳
                    if (ce.ChkFMEdStatusScrap(row["ScrapLotNum"].ToString().Trim()))
                    {
                        ewnas.DeleteScrapWIPLog(row["ID"].ToString());
                    }
                    //再檢查是否有開報廢單
                    else if (ce.ChkFMEdScrap(row["ScrapLotNum"].ToString()))
                    {
                        ewnas.DeleteScrapWIPLog(row["ID"].ToString());
                    }
                    else
                    {
                        result.ImportRow(row);
                    }
                }
            }
            result.Columns.RemoveAt(0);
            result.Columns[0].ColumnName="製令單號";
            result.Columns[1].ColumnName="單據日期";
            result.Columns[2].ColumnName="料號";
            result.Columns[3].ColumnName="版序";
            result.Columns[4].ColumnName="訂單類型";
            result.Columns[5].ColumnName="製令總Pcs數";
            result.Columns[6].ColumnName="期望繳庫日";
            result.Columns[7].ColumnName="審核人員";
            result.Columns[8].ColumnName="提出製程";
            result.Columns[9].ColumnName="代碼";
            result.Columns[10].ColumnName="批號";
            result.Columns[11].ColumnName="狀態";
            result.Columns[12].ColumnName="階段名稱";
            result.Columns[13].ColumnName="型狀";
            result.Columns[14].ColumnName="排版";
            result.Columns[15].ColumnName="預報廢數量";
            result.Columns[16].ColumnName="批量種類";
            try
            {
                SaveFile = Path.GetTempPath() + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") +
                    "清單.xls";
                DataTable[] resultTB = new DataTable[] { result };
                string[] strSheet = new string[] { "防焊現帳預報廢未增帳清單" };
                DataTableToExcel(resultTB, strSheet, SaveFile);
            }
            catch (Exception ex)
            {
                InsertLog(DateTime.Now.ToString(datetimeFormat) + "  ChkScrapWIPLog()" +
                    ex.Message + "\n\r");
            }
            finally
            {
                result.Dispose();
                ScrapWIPLog.Dispose();
            }
            SendMail("sm4@ewpcb.com.tw", "防焊現帳預報廢未增帳清單", "chkissuescrapwip@ewpcb.com.tw",
                    DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + " 防焊現帳預報廢未增帳清單",
                    "防焊現帳預報廢未增帳清單，請詳閱附件。" + "<br/>" + "<br/>" +
                    "-----此封郵件由系統所寄出，請勿直接回覆！-----", SaveFile);
        }

        /// <summary>
        /// 稽核輔助系統生產日報表
        /// 稽核時間為前一天的早、晚班
        /// </summary>
        private void ChkProductDailyReport()
        {
            ConnEWNAS ewnas = new ConnEWNAS();
            var fromDate = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 08:10:00");
            var endDate = DateTime.Now.ToString("yyyy-MM-dd 08:10:00");
            var strResult = string.Empty;
            var xlsResult = new DataTable[6];
            try
            {
                #region 防焊顯影
                var ResultLF = ewnas.ChkDevelopmentProductDailyReportLF(fromDate, endDate);
                xlsResult[0] = ResultLF.Copy();
                strResult += string.Format("稽核區間：{0} ~ {1}<p>", fromDate, endDate);
                strResult += string.Format("{0}\t{1}\t{2}\t{3}\t\t{4}\t\t{5}\t{6}\t{7}\t{8}\t{9}\t\t\t{10}<br>",
                    "項次", "姓名", "部門", "批號", "料號", "層別", "工序", "機台", "數量", "開始時間", "結束時間");
                if (ResultLF.Rows.Count != 0)
                {
                    foreach (DataRow row in ResultLF.Rows)
                    {
                        strResult += string.Format("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}\t{10}<br>",
                                row["Item"].ToString(),
                                row["empname"].ToString(),
                                row["departname"].ToString(),
                                row["lotnum"].ToString(),
                                row["partnum"].ToString(),
                                row["layername"].ToString(),
                                row["process"].ToString(),
                                row["machineno"].ToString(),
                                row["workqnty"].ToString(),
                                row["starttime"].ToString(),
                                row["endtime"].ToString());
                    }
                    strResult += "<p>上述筆數為Ewproject有申報但輔助系統防焊顯影生產日報查無的紀錄！<p>" +
                        "以上為防焊顯影生產日報漏申報稽核結果。<p>";
                }
                else
                {
                    strResult += "<br>防焊顯影生產日報查無漏申報紀錄！<p>";
                }
                ResultLF.Dispose();
                SendMail("sm4@ewpcb.com.tw", "防焊顯影生產日報稽核結果", "chkdeclarelf@ewpcb.com.tw",
                    DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + " 防焊顯影生產日報稽核結果！",
                    "<pre>" + strResult + "</pre>-----此封郵件由系統所寄出，請勿直接回覆！-----", null);
                #endregion

                #region 乾膜顯影
                var ResultFF = ewnas.ChkDevelopmentProductDailyReportFF(fromDate, endDate);
                xlsResult[1]=ResultFF.Copy();
                strResult = string.Empty;
                strResult += string.Format("稽核區間：{0} ~ {1}<p>", fromDate, endDate);
                strResult += string.Format("{0}\t{1}\t{2}\t{3}\t\t{4}\t\t{5}\t{6}\t{7}\t{8}\t{9}\t\t\t{10}<br>",
                    "項次", "姓名", "部門", "批號", "料號", "層別", "工序", "機台", "數量", "開始時間", "結束時間");
                if (ResultFF.Rows.Count != 0)
                {
                    foreach (DataRow row in ResultFF.Rows)
                    {
                        strResult += string.Format("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}\t{10}<br>",
                                row["Item"].ToString(),
                                row["empname"].ToString(),
                                row["departname"].ToString(),
                                row["lotnum"].ToString(),
                                row["partnum"].ToString(),
                                row["layername"].ToString(),
                                row["process"].ToString(),
                                row["machineno"].ToString(),
                                row["workqnty"].ToString(),
                                row["starttime"].ToString(),
                                row["endtime"].ToString());
                    }
                    strResult += "<p>上述筆數為Ewproject有申報但輔助系統乾膜顯影生產日報查無的紀錄！<p>" +
                        "以上為乾膜顯影生產日報漏申報稽核結果。<p>";
                }
                else
                {
                    strResult += "<br>乾膜顯影生產日報查無漏申報紀錄！<p>";
                }
                ResultFF.Dispose();
                SendMail("sm4@ewpcb.com.tw", "乾膜顯影生產日報稽核結果", "chkdeclareff@ewpcb.com.tw",
                    DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + " 乾膜顯影生產日報稽核結果！",
                    "<pre>" + strResult + "</pre>-----此封郵件由系統所寄出，請勿直接回覆！-----", null);
                #endregion

                #region 乾膜壓膜
                var Result壓膜 = ewnas.Chk乾膜壓膜生產日報表(fromDate, endDate);
                xlsResult[2]=Result壓膜.Copy();
                strResult = string.Empty;
                strResult += string.Format("稽核區間：{0} ~ {1}<p>", fromDate, endDate);
                strResult += string.Format("{0}\t{1}\t{2}\t{3}\t\t{4}\t\t{5}\t{6}\t{7}\t{8}\t{9}\t\t\t{10}<br>",
                    "項次", "姓名", "部門", "批號", "料號", "層別", "工序", "機台", "數量", "開始時間", "結束時間");
                if (Result壓膜.Rows.Count != 0)
                {
                    foreach (DataRow row in Result壓膜.Rows)
                    {
                        strResult += string.Format("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}\t{10}<br>",
                                row["Item"].ToString(),
                                row["empname"].ToString(),
                                row["departname"].ToString(),
                                row["lotnum"].ToString(),
                                row["partnum"].ToString(),
                                row["layername"].ToString(),
                                row["process"].ToString(),
                                row["machineno"].ToString(),
                                row["workqnty"].ToString(),
                                row["starttime"].ToString(),
                                row["endtime"].ToString());
                    }
                    strResult += "<p>上述筆數為Ewproject有申報但輔助系統乾膜壓膜日報查無的紀錄！<p>" +
                        "以上為乾膜壓膜生產日報漏申報稽核結果。<p>";
                }
                else
                {
                    strResult += "<br>乾膜壓膜生產日報查無漏申報紀錄！<p>";
                }
                Result壓膜.Dispose();
                SendMail("sm4@ewpcb.com.tw", "乾膜壓膜生產日報稽核結果", "chkdeclareff@ewpcb.com.tw",
                    DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + " 乾膜壓膜生產日報稽核結果！",
                    "<pre>" + strResult + "</pre>-----此封郵件由系統所寄出，請勿直接回覆！-----", null);
                #endregion

                #region 乾膜AOI
                var ResultAOI = ewnas.Chk乾膜AOI檢查日報表(fromDate, endDate);
                xlsResult[3]=ResultAOI.Copy();
                strResult = string.Empty;
                strResult += string.Format("稽核區間：{0} ~ {1}<p>", fromDate, endDate);
                strResult += string.Format("{0}\t{1}\t{2}\t{3}\t\t{4}\t\t{5}\t{6}\t{7}\t{8}\t{9}\t\t\t{10}<br>",
                    "項次", "姓名", "部門", "批號", "料號", "層別", "工序", "機台", "數量", "開始時間", "結束時間");
                if (ResultAOI.Rows.Count != 0)
                {
                    foreach (DataRow row in ResultAOI.Rows)
                    {
                        strResult += string.Format("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}\t{10}<br>",
                                row["Item"].ToString(),
                                row["empname"].ToString(),
                                row["departname"].ToString(),
                                row["lotnum"].ToString(),
                                row["partnum"].ToString(),
                                row["layername"].ToString(),
                                row["process"].ToString(),
                                row["machineno"].ToString(),
                                row["workqnty"].ToString(),
                                row["starttime"].ToString(),
                                row["endtime"].ToString());
                    }
                    strResult += "<p>上述筆數為Ewproject有申報但輔助系統乾膜AOI檢查日報查無的紀錄！<p>" +
                        "以上為乾膜AOI檢查日報漏申報稽核結果。<p>";
                }
                else
                {
                    strResult += "<br>乾膜AOI檢查日報查無漏申報紀錄！<p>";
                }
                ResultAOI.Dispose();
                SendMail("sm4@ewpcb.com.tw", "乾膜AOI檢查日報稽核結果", "chkdeclareff@ewpcb.com.tw",
                    DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + " 乾膜AOI檢查日報稽核結果！",
                    "<pre>" + strResult + "</pre>-----此封郵件由系統所寄出，請勿直接回覆！-----", null);
                #endregion

                #region 成型V-CUT生產量測
                var ResultVCUT = ewnas.ChkVCutProductCheckRepay(fromDate, endDate);
                xlsResult[4]=ResultVCUT.Copy();
                strResult = string.Empty;
                strResult += string.Format("稽核區間：{0} ~ {1}<p>", fromDate, endDate);
                strResult += string.Format("{0}\t{1}\t{2}\t{3}\t\t{4}\t\t{5}\t{6}\t{7}\t{8}\t{9}\t\t\t{10}<br>",
                    "項次", "姓名", "部門", "批號", "料號", "層別", "工序", "機台", "數量", "開始時間", "結束時間");
                if (ResultVCUT.Rows.Count != 0)
                {
                    foreach (DataRow row in ResultVCUT.Rows)
                    {
                        strResult += string.Format("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}\t{10}<br>",
                                row["Item"].ToString(),
                                row["empname"].ToString(),
                                row["departname"].ToString(),
                                row["lotnum"].ToString(),
                                row["partnum"].ToString(),
                                row["layername"].ToString(),
                                row["process"].ToString(),
                                row["machineno"].ToString(),
                                row["workqnty"].ToString(),
                                row["starttime"].ToString(),
                                row["endtime"].ToString());
                    }
                    strResult += "<p>上述筆數為Ewproject有申報但輔助系統成型V-CUT生產/量測日報查無的紀錄！<p>" +
                        "以上為成型V-CUT生產/量測日報漏申報稽核結果。<p>";
                }
                else
                {
                    strResult += "<br>成型V-CUT生產/量測日報查無漏申報紀錄！<p>";
                }
                ResultVCUT.Dispose();
                SendMail("sm4@ewpcb.com.tw", "成型V-CUT生產/量測日報稽核結果", "chkdeclarecf@ewpcb.com.tw",
                    DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + " 成型V-CUT生產/量測日報稽核結果！",
                    "<pre>" + strResult + "</pre>-----此封郵件由系統所寄出，請勿直接回覆！-----", null);
                #endregion

                #region 壓合PP裁切自主檢查
                var ResultPPCUT = ewnas.ChkPPCutChkReport(fromDate, endDate);
                xlsResult[5]=ResultPPCUT.Copy();
                strResult = string.Empty;
                strResult += string.Format("稽核區間：{0} ~ {1}<p>", fromDate, endDate);
                strResult += string.Format("{0}\t{1}\t{2}\t{3}\t\t{4}\t\t{5}\t{6}\t{7}\t{8}\t{9}\t\t\t{10}<br>",
                    "項次", "姓名", "部門", "批號", "料號", "層別", "工序", "機台", "數量", "開始時間", "結束時間");
                if (ResultPPCUT.Rows.Count != 0)
                {
                    foreach (DataRow row in ResultPPCUT.Rows)
                    {
                        strResult += string.Format("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}\t{10}<br>",
                                row["Item"].ToString(),
                                row["empname"].ToString(),
                                row["departname"].ToString(),
                                row["lotnum"].ToString(),
                                row["partnum"].ToString(),
                                row["layername"].ToString(),
                                row["process"].ToString(),
                                row["machineno"].ToString(),
                                row["workqnty"].ToString(),
                                row["starttime"].ToString(),
                                row["endtime"].ToString());
                    }
                    strResult += "<p>上述筆數為Ewproject有申報但輔助系統壓合PP裁切自主檢查日報查無的紀錄！<p>" +
                        "以上為壓合PP裁切自主檢查日報漏申報稽核結果。<p>";
                }
                else
                {
                    strResult += "<br>壓合PP裁切自主檢查日報查無漏申報紀錄！<p>";
                }
                ResultPPCUT.Dispose();
                SendMail("sm4@ewpcb.com.tw", "壓合PP裁切自主檢查日報稽核結果", "chkdeclareel@ewpcb.com.tw",
                    DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + " 壓合PP裁切自主檢查日報稽核結果！",
                    "<pre>" + strResult + "</pre>-----此封郵件由系統所寄出，請勿直接回覆！-----", null);
                #endregion

                #region 單獨滙整成一份EXCEL檔，寄給廠長 david@ewpcb.com.tw 和製研 me3@ewpcb.com.tw
                //xlsResult[0] = ewnas.ChkDevelopmentProductDailyReportLF(fromDate, endDate);
                //xlsResult[1] = ewnas.ChkDevelopmentProductDailyReportFF(fromDate, endDate);
                //xlsResult[2] = ewnas.Chk乾膜壓膜生產日報表(fromDate, endDate);
                //xlsResult[3] = ewnas.Chk乾膜AOI檢查日報表(fromDate, endDate);
                //xlsResult[4] = ewnas.ChkVCutProductCheckRepay(fromDate, endDate);
                //xlsResult[5] = ewnas.ChkPPCutChkReport(fromDate, endDate);
                SaveFile = Path.GetTempPath() + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") +
                    ".xls";
                var sheetName = new string[] { "防焊顯影", "乾膜顯影", "乾膜壓膜", "乾膜AOI", "成型V-CUT", "壓合PP裁切" };
                DataTableToExcel(xlsResult, sheetName, SaveFile);
                SendMail("sm4@ewpcb.com.tw", "輔助系統表單稽核結果", "chkman@ewpcb.com.tw",
                    DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + " 輔助系統表單稽核結果！",
                    "輔助系統表單稽核結果，請詳閱附件。<br/><br/>-----此封郵件由系統所寄出，請勿直接回覆！-----", SaveFile);
                #endregion
            }
            catch (Exception ex)
            {
                InsertLog("ChkProductDailyReport() " + ex.Message);
            }
        }

        /// <summary>
        /// 稽核每日品保IQC檢驗申報表單申報記錄
        /// </summary>
        private void ChkIQC每日量測申報紀錄()
        {
            ConnEWNAS ewnas = new ConnEWNAS();
            var Copper = ewnas.ChkIQC一二銅量測申報();
            var Slice = ewnas.ChkIQC切片量測申報();
            var BaseBoard = ewnas.ChkIQC基板銅箔量測申報();
            var CutCopper = ewnas.ChkIQC磨刷減銅量測申報();
            var xlsResult = new DataTable[] { Copper, Slice, BaseBoard, CutCopper };
            SaveFile = Path.GetTempPath() + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + ".xls";
            var sheetName = new string[] { "一、二銅", "切片", "基板銅箔", "磨刷減銅" };
            DataTableToExcel(xlsResult, sheetName, SaveFile);
            SendMail("sm4@ewpcb.com.tw", "輔助系統IQC表單稽核結果", "chkdeclareqc@ewpcb.com.tw",
                DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + " 輔助系統IQC表單稽核結果！",
                "輔助系統IQC表單稽核結果，請詳閱附件。<br/><br/>-----此封郵件由系統所寄出，請勿直接回覆！-----", SaveFile);
            Copper.Dispose();
            Slice.Dispose();
            BaseBoard.Dispose();
            CutCopper.Dispose();
        }

        /// <summary>
        /// 查詢倉庫不足安全庫存量物料
        /// </summary>
        private void ChkDepotStock()
        {
            ConnERP erp = new ConnERP();
            var Result = erp.CheckDepotStock();
            var TB = new DataTable[] { Result };
            var Sheet = new string[] { DateTime.Now.ToString("yyyy-MM-dd_HHmm") };
            SaveFile = Path.GetTempPath() + DateTime.Now.ToString("yyyy-MM-dd_HHmm") + "清單.xls";
            DataTableToExcel(TB, Sheet, SaveFile);
            SendMail("sm4@ewpcb.com.tw", "倉庫不足安全庫存量物料結果", "depotstock@ewpcb.com.tw",
                DateTime.Now.ToString("yyyy-MM-dd_HH:mm") + " 倉庫不足安全庫存量物料！",
                "倉庫不足安全庫存量物料，請詳閱附件。<br/><br/>-----此封郵件由系統所寄出，請勿直接回覆！-----", SaveFile);
            Result.Dispose();
        }

        /// <summary>
        /// 檢查V-CUT跳刀料號目前在那一站
        /// </summary>
        private void ChkVCUT_Jump()
        {
            CheckVCUT_Jump chk = new CheckVCUT_Jump();
            var result = chk.GetResult();
            if (result.Rows.Count > 0)
            {
                var strResult = string.Empty;
                CheckVCUT_JumpLog log = new CheckVCUT_JumpLog();

                #region 測試用資料
                //strResult += string.Format("{0}<br/>{1}<br/>{2}<br/>{3}<br/>{4}<br/>{5}<br/>{6}<br/>{7}<br/>{8}<br/>",
                //        "批號：L708010006 ",
                //        "料號：M09-2049-CX", 
                //        "版序：01AX", 
                //        "層別：L0~0", 
                //        "片型：WPnl", 
                //        "數量：7", 
                //        "重工：0",
                //        "目前途程：220 鑽孔1st",
                //        "下站途程：410 防焊(S/M)");
                #endregion

                foreach (DataRow row in result.Rows)
                {
                    if (!log.Check(row["批號"].ToString().Trim(), row["目前途程"].ToString().Substring(0, 4),
                        row["下站途程"].ToString().Substring(0, 4)))
                    {
                        log.Insert(row["目前途程"].ToString().Substring(0, 4), row["批號"].ToString().Trim(),
                            row["料號"].ToString().Trim(), row["版序"].ToString().Trim(), row["層別"].ToString().Trim(),
                            row["片型"].ToString().Trim(), row["數量"].ToString().Trim(), row["重工"].ToString().Trim(),
                            row["下站途程"].ToString().Substring(0, 4));
                        strResult = string.Format("{0}<br/>{1}<br/>{2}<br/>{3}<br/>{4}<br/>{5}<br/>{6}<br/>{7}<br/>{8}<br/>",
                            "批號：" + row["批號"].ToString(),
                            "料號：" + row["料號"].ToString(),
                            "版序：" + row["版序"].ToString(),
                            "層別：" + row["層別"].ToString(),
                            "片型：" + row["片型"].ToString(),
                            "數量：" + row["數量"].ToString(),
                            "重工：" + row["重工"].ToString(),
                            "目前途程：" + row["目前途程"].ToString(),
                            "下站途程：" + row["下站途程"].ToString());
                        SendMail("sm4@ewpcb.com.tw", "V-CUT跳刀料號進站通知", "checkvcutjump@ewpcb.com.tw",
                            DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + " V-CUT跳刀料號進站通知！",
                            strResult + "<br/><br/>-----此封郵件由系統所寄出，請勿直接回覆！-----", null);
                    }
                }
            }
        }

        /// <summary>
        /// 檢查所有未交貨完畢製令單的料號特殊注意事項
        /// </summary>
        private void ChkTracePartNum()
        {
            /*
             * 2017/01/08 CoreyChen Modify
             * 總經理指示，將需注意的特殊料號事項合併，不要依個別製令單號顯示！
             * 因為若依製令單號顯示，重覆料號太多，故修正成依料號顯示並將製令單總PCS數加總。
             */
            var Result = new DataTable();
            //Result.Columns.Add("製令單號");
            //Result.Columns.Add("作業時間");
            //Result.Columns.Add("訂單種類");
            //Result.Columns.Add("批量種類");
            Result.Columns.Add("料號");
            Result.Columns.Add("版序");
            Result.Columns.Add("製令總PCS數");
            Result.Columns.Add("特殊注意事項");
            CheckFMEdIssueNote chk = new CheckFMEdIssueNote();
            var srcData = chk.GetAllFMEdIssue();
            foreach (DataRow row in srcData.Rows)
            {
                var SpecialData = chk.GetTeacePartNum(row["料號"].ToString().Substring(0, 7));
                if (SpecialData.Rows.Count > 0)
                {
                    DataRow NewRow = Result.NewRow();
                    //NewRow["製令單號"] = row["製令單號"].ToString();
                    //NewRow["作業時間"] = row["作業時間"].ToString();
                    //NewRow["訂單種類"] = row["訂單種類"].ToString();
                    //NewRow["批量種類"] = row["批量種類"].ToString();
                    NewRow["料號"] = row["料號"].ToString();
                    NewRow["版序"] = row["版序"].ToString();
                    NewRow["製令總PCS數"] = row["製令總PCS數"].ToString();
                    for (int i = 0; i < SpecialData.Rows.Count; i++)
                    {
                        NewRow["特殊注意事項"] += (i + 1) + ". " + SpecialData.Rows[i]["特殊事項"].ToString().Trim() +
                            " ";
                    }
                    Result.Rows.Add(NewRow);
                }
            }
            SaveFile = Path.GetTempPath() + @"\SpecialPartNum.xls";
            var srcTable = new DataTable[] { Result };
            var srcSheetName = new string[] { "製令單特殊料號注意事項" };
            DataTableToExcel(srcTable, srcSheetName, SaveFile);
            SendMail("sm4@ewpcb.com.tw", "製令單特殊料號注意事項通知", "chkfmedissuenote@ewpcb.com.tw", 
                DateTime.Now.ToString("yyyy-MM-dd_HHmm") + " 製令單特殊料號注意事項通知！",
                "製令單特殊料號注意事項通知，請詳閱附件。<br/><br/>-----此封郵件由系統所寄出，請勿直接回覆！-----", SaveFile);
        }

        /// <summary>
        /// 檢查每月工程部NAS鑽孔程式資料夾裡的料號是有含二次鑽程式
        /// 若有含就要去檢查目前製程現帳中該料號的途程是否有420這一站！
        /// 若沒有就要發信通知工程、生管、廠長，每天早上和下午各檢查一次。
        /// </summary>
        private void ChkDrill2ndProcByMonth()
        {
            var Result = string.Format("{0}　　　　　{1}　　　　　　{2}　　　{3}　　　{4}<br/>",
                "批號", "料號", "版序", "數量", "目前途程");
            //開發測試時使用
            //var srcDirectory = @"\\n5200xxx\鑽孔與成型二\" + DateTime.Now.ToString("yyyy-MM") + @"\";
            var srcDirectory = @"D:\鑽孔與成型二\" + DateTime.Now.ToString("yyyy-MM") + @"\";
            try
            {
                DirectoryInfo directoryMain = new DirectoryInfo(srcDirectory);
                var directorySub = directoryMain.GetDirectories();
                foreach (DirectoryInfo GetDirSubInfo in directorySub)
                {
                    var PartNumDir = GetDirSubInfo.GetDirectories();
                    foreach (DirectoryInfo PartNumDirSub in PartNumDir)
                    {
                        var strFilePath = PartNumDirSub.FullName + @"\" + PartNumDirSub.Name + ".2nd";
                        if (File.Exists(strFilePath))
                        {
                            var PartNum = PartNumDirSub.Name.Substring(0, 11);
                            var Revision = PartNumDirSub.Name.Substring(12, 3);
                            ConnERP connERP = new ConnERP();
                            var tempResult = connERP.CheckFMEdProcDrill2nd(PartNum, Revision);
                            if (tempResult.Rows.Count != 0)
                            {
                                foreach (DataRow row in tempResult.Rows)
                                {
                                    Result += string.Format("{0}　　{1}　　{2}　　　{3}　　　　　{4}<br/>",
                                        row["LotNum"].ToString(),
                                        row["PartNum"].ToString(),
                                        row["Revision"].ToString(),
                                        row["Qnty"].ToString(),
                                        row["ProcCode"].ToString());
                                }
                            }
                        }
                    }
                }
                SendMail("sm4@ewpcb.com.tw", "二次鑽料號工單漏420途程通知", "drill2nd@ewpcb.com.tw",
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm") + " 二次鑽料號工單漏420途程通知！",
                    Result + "<br/><br/>-----此封郵件由系統所寄出，請勿直接回覆！-----", null);
            }
            catch (Exception ex)
            {
                InsertLog("ChkDrill2ndProcByMonth()-" + ex.Message);
            }
        }

        /// <summary>
        /// 檢查工程承認書製作料號清單
        /// 若未押指定完成日期，則在傑偲製程現帳檢查相關工單中，只要任一張工單已進到650測試站，
        /// 尚未結案且未發過通知信的的承認書單據需寄出MAIL通知相關人員。
        /// </summary>
        private void CheckAcknowledgmentIn650()
        {
            ConnEWNAS ewnas = new ConnEWNAS();
            var SrcTB = ewnas.ChkAcknowledgmentIn650();
            if (SrcTB.Rows.Count > 0)
            {
                var Result = string.Empty;
                foreach (DataRow row in SrcTB.Rows)
                {
                    Result += string.Format("<br/>{0}<br/>{1}<br/>{2}<br/>{3}<br/>{4}<br/>{5}<br/>{6}<br/>{7}<br/>" +
                        "======================================</br>",
                        "單據號碼：" + row["單據號碼"].ToString(),
                        "料號：" + row["料號"].ToString(),
                        "建立時間：" + row["建立時間"].ToString(),
                        "建立人員：" + row["建立人員"].ToString(),
                        "出貨報告完成時間：" + row["出貨報告完成時間"].ToString(),
                        "出貨報告完成人員：" + row["出貨報告完成人員"].ToString(),
                        "承認書完成時間：" + row["承認書完成時間"].ToString(),
                        "承認書完成人員：" + row["承認書完成人員"].ToString());
                    ewnas.UpdateAcknowledgmentSendMail(row["單據號碼"].ToString());
                }
                try
                {
                    SendMail("sm4@ewpcb.com.tw", "承認書製作通知", "acknowledgment@ewpcb.com.tw",
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm") + " 料號承認書請開始製作，料號已進測試站通知！",
                    Result + "<br/><br/>-----此封郵件由系統所寄出，請勿直接回覆！-----", null);
                }
                catch (Exception ex)
                {
                    InsertLog("CheckAcknowledgmentIn650()-" + ex.Message);
                }
            }
        }

        /// <summary>
        /// 檢查工程承認書製作料號清單
        /// 若有押指定完成日期，則要在完成日的前3天就通知
        /// 尚未結案且未發過通知信的的承認書單據需寄出MAIL通知相關人員。
        /// </summary>
        private void CheckAcknowledgmentAppointDate()
        {
            ConnEWNAS ewnas = new ConnEWNAS();
            var SrcTB = ewnas.ChkAcknowledgmentAppointDate();
            if (SrcTB.Rows.Count > 0)
            {
                var Result = string.Empty;
                foreach (DataRow row in SrcTB.Rows)
                {
                    var AppointDate = Convert.ToDateTime(row["指定完成日期"]);
                    var DifferenceDay = AppointDate - DateTime.Now.Date;
                    if (DifferenceDay.TotalDays <= 3)
                    {
                        Result += string.Format("<br/>{0}<br/>{1}<br/>{2}<br/>{3}<br/>{4}<br/>{5}<br/>{6}<br/>{7}<br/>" +
                            "{8}<br/>======================================<br/>",
                            "單據號碼：" + row["單據號碼"].ToString(),
                            "料號：" + row["料號"].ToString(),
                            "建立時間：" + row["建立時間"].ToString(),
                            "建立人員：" + row["建立人員"].ToString(),
                            "指定完成日期：" + row["指定完成日期"].ToString(),
                            "出貨報告完成時間：" + row["出貨報告完成時間"].ToString(),
                            "出貨報告完成人員：" + row["出貨報告完成人員"].ToString(),
                            "承認書完成時間：" + row["承認書完成時間"].ToString(),
                            "承認書完成人員：" + row["承認書完成人員"].ToString());
                        ewnas.UpdateAcknowledgmentSendMail(row["單據號碼"].ToString());
                    }
                }
                if (!string.IsNullOrWhiteSpace(Result))
                {
                    try
                    {
                        SendMail("sm4@ewpcb.com.tw", "承認書製作通知", "acknowledgment@ewpcb.com.tw",
                        DateTime.Now.ToString("yyyy-MM-dd HH:mm") + " 料號承認書請開始製作，指定完成日即將到期！",
                        Result + "<br/><br/>-----此封郵件由系統所寄出，請勿直接回覆！-----", null);
                    }
                    catch (Exception ex)
                    {
                        InsertLog("CheckAcknowledgmentAppointDate()-" + ex.Message);
                    }
                }
            }
        }

        /// <summary>
        /// 取得特休假即將到期的在職人員清冊
        /// </summary>
        private void HPSdGetSpecialBreakDay()
        {
            ConnERP erp = new ConnERP();
            var SrcData = erp.HPSdGetSpecialBreakDay();
            if (SrcData.Rows.Count > 0)
            {
                try
                {
                    var TB = new DataTable[] { SrcData };
                    var Sheet = new string[] { "特休日即將到期人員清單" };
                    SaveFile = Path.GetTempPath() + DateTime.Now.ToString("yyyy-MM-dd") + "特休到期清單.xls";
                    DataTableToExcel(TB, Sheet, SaveFile);
                    SendMail("sm4@ewpcb.com.tw", "特休假到期通知", "specialbreakday@ewpcb.com.tw",
                        DateTime.Now.ToString("yyyy-MM-dd_HH:mm") + " 特休假即將到期人員清單！",
                        "特休假即將到期人員清單，請詳閱附件。<br/><br/>-----此封郵件由系統所寄出，請勿直接回覆！-----", SaveFile);
                    SrcData.Dispose();
                }
                catch (Exception ex)
                {
                    InsertLog("HPSdGetSpecialBreakDay()-" + ex.Message);
                }
            }
        }

        /// <summary>
        /// 將資料表轉成Eecel檔案
        /// 若轉出並存檔成功，傳回true
        /// </summary>
        /// <param name="Source">要轉出的資料表陣列</param>
        /// <param name="SheetName">工作表名陣列</param>
        /// <param name="SavePath">存檔路徑及檔名</param>
        /// <param name="ExcelRevi">要存檔的Excel版本，2007以上為true</param>
        private bool DataTableToExcel(DataTable[] Source, string[] SheetName, string SavePath)
        {
            var cc = DateTime.Now.ToString();
            bool result = false;
            IWorkbook workbook;
            ISheet Sheet;
            workbook = new HSSFWorkbook();
            HSSFCellStyle csTitle = (HSSFCellStyle)workbook.CreateCellStyle();
            HSSFCellStyle csCell = (HSSFCellStyle)workbook.CreateCellStyle();
            HSSFFont fontTitle = (HSSFFont)workbook.CreateFont();
            HSSFFont fontCell = (HSSFFont)workbook.CreateFont();
            fontTitle.FontName = "微軟正黑體";
            fontTitle.FontHeightInPoints = 11;
            fontTitle.Boldweight = (short)FontBoldWeight.Bold;
            fontCell.FontName = "微軟正黑體";
            fontCell.FontHeightInPoints = 11;
            csTitle.SetFont(fontTitle);
            csCell.SetFont(fontCell);
            for (int AllSheet = 0; AllSheet < Source.Length; AllSheet++)
            {
                if (!string.IsNullOrEmpty(SheetName[AllSheet]))
                {
                    Sheet = workbook.CreateSheet(SheetName[AllSheet]);
                }
                else
                {
                    Sheet = workbook.CreateSheet("Sheet" + AllSheet);
                }
                Sheet.CreateRow(0);
                for (int i = 0; i < Source[AllSheet].Columns.Count; i++)
                {
                    Sheet.GetRow(0).CreateCell(i).SetCellValue(Source[AllSheet].Columns[i].ColumnName);
                    HSSFCell cell = (HSSFCell)Sheet.GetRow(0).GetCell(i);
                    cell.CellStyle = csTitle;
                }
                for (int i = 0; i < Source[AllSheet].Rows.Count; i++)
                {
                    Sheet.CreateRow(i + 1);
                    for (int x = 0; x < Source[AllSheet].Columns.Count; x++)
                    {
                        Sheet.GetRow(i + 1).CreateCell(x).SetCellValue(Source[AllSheet].Rows[i][x].ToString());
                        HSSFCell cell = (HSSFCell)Sheet.GetRow(i + 1).GetCell(x);
                        cell.CellStyle = csCell;
                    }
                }
                for (int i = 0; i < Source[AllSheet].Columns.Count; i++)
                {
                    Sheet.AutoSizeColumn(i);
                }
            }
            FileStream SaveFile = new FileStream(SavePath, FileMode.Create);
            try
            {
                workbook.Write(SaveFile);
                result = true;
            }
            catch (Exception ex)
            {
                InsertLog(DateTime.Now.ToString(datetimeFormat) + "  " + ex.Message + "\r\n");
            }
            finally
            {
                workbook.Close();
                SaveFile.Flush();
                SaveFile.Close();
                SaveFile.Dispose();
            }
            return result;
        }

        /// <summary>
        /// 寄出電子郵件
        /// </summary>
        /// <param name="from">寄件者地址</param>
        /// <param name="display">寄件者名稱</param>
        /// <param name="to">收件人地址</param>
        /// <param name="sub">郵件主旨</param>
        /// <param name="body">郵件內容</param>
        /// <param name="att">郵件附件，若無附件可為null</param>
        private void SendMail(string from, string display, string to, string sub, string body, string att)
        {
            //建立寄件者地址與名稱
            MailAddress ReceiverAddress = new MailAddress(from, display);
            //建立收件者地址
            MailAddress SendAddress = new MailAddress(to);
            //建立E-MAIL相關設定與訊息
            using (MailMessage SendMail = new MailMessage(ReceiverAddress, SendAddress))
            {
                //Mail以HTML格式寄送
                SendMail.IsBodyHtml = true;

                /*
                ##### 2016/12/08 #####
                設定信件標頭走Base64傳輸，避免中文檔名附件變成亂碼
                */
                SendMail.Headers.Add("Content-Transfer-Encoding", "base64");

                //設定信件主旨編碼為UTF8
                SendMail.SubjectEncoding = Encoding.UTF8;

                //設定信件內容編碼為UTF8
                SendMail.BodyEncoding = Encoding.UTF8;
                
                //設定信件優先權為普通
                SendMail.Priority = MailPriority.Normal;
                SendMail.Subject = sub;//主旨
                SendMail.Body = body;//內容

                if (!string.IsNullOrEmpty(att))
                {
                    //建立附加檔案
                    Attachment attachment = new Attachment(att);
                    SendMail.Attachments.Add(attachment);//加上附件檔案

                    //建立一個信件通訊並設定郵件主機地址與通訊埠號
                    using (SmtpClient MySmtp = new SmtpClient("ms1.ewpcb.com.tw", 25))
                    {
                        //設定寄件者的帳號與密碼
                        MySmtp.Credentials = new NetworkCredential("sm4", "sm4@ew");
                        try
                        {
                            MySmtp.Send(SendMail);
                            //InsertLog(sub + "   Send Mail OK!");
                        }
                        catch (Exception ex)
                        {
                            InsertLog(sub + " " + ex.Message + "\r\n");
                        }
                        finally
                        {
                            attachment.Dispose();
                            SendMail.Dispose();
                        }
                    }
                }
                else
                {
                    using (SmtpClient MySmtp = new SmtpClient("ms1.ewpcb.com.tw", 25))
                    {
                        MySmtp.Credentials = new NetworkCredential("sm4", "sm4@ew");
                        try
                        {
                            MySmtp.Send(SendMail);
                            //InsertLog(sub + "   Send Mail OK!");
                        }
                        catch (Exception ex)
                        {
                            InsertLog(sub + " " + ex.Message + "\r\n");
                        }
                        finally
                        {
                            SendMail.Dispose();
                        }
                    }
                }
            }
            if (!string.IsNullOrEmpty(att))
            {
                File.Delete(SaveFile);//刪除存放在系統個人暂存區的檔案
            }
        }
    }
}
