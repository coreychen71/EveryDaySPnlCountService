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
        private Timer timer;
        //private DateTime setTime;
        private DateTime nowTime;
        //設定間隔時間為1分鐘
        private double timerInterval = 60 * 1000;
        private string datetimeFormat = "yyyy-MM-dd HH:mm:ss";
        private string SaveFile="";
        private string LogPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + 
            "\\EveryDaySPnlCountLog.txt";
        private StreamWriter writerResult;
        private static StreamWriter writerLog;

        public MainMethod()
        {
            timer = new Timer();
            timer.Interval = timerInterval;
            timer.AutoReset = false;
            timer.Enabled = false;
            timer.Elapsed += Timer_Elapsed;
            //此方式為每次寫入時，持續寫入，不會覆蓋原本內容
            writerLog = File.AppendText(LogPath);
        }

        /// <summary>
        /// 寫入Log
        /// </summary>
        /// <param name="Msg">要寫入的訊息</param>
        public static void InsertLog(string Msg)
        {
            writerLog.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + Msg + "\r\n");
            writerLog.Flush();
        }

        private void Timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            timer.Stop();
            //每日06: 30進行SPnl數統計(TXT)
            if (CheckTime("06:30:00", "06:30:59"))
            {
                SPnlCountRun();
            }
            //每日07:00進行品保客訴待處理(未逾期)清單(Excel)
            else if (CheckTime("07:00:00", "07:00:59"))
            {
                GetEveryDayCustomerComplaint();
            }
            //每日16:00進行待修設備清單(Excel)
            else if (CheckTime("16:00:00", "16:00:59"))
            {
                EquipMaintainRun();
            }
            //每日00:10進行每日製令稽核現帳預報廢清單(Excel)
            else if (CheckTime("00:10:00", "00:10:59"))
            {
                ChkIssueAndScrapWIP();
            }
            //每日00:20進行鑽孔申報比對驗孔紀錄(TXT)
            else if (CheckTime("00:20:00", "00:20:59"))
            {
                ChkDrillHole();
            }
            //每日00: 30進行製令單料號檢查是否有特殊油墨與樹脂需求(TXT)
            else if (CheckTime("00:30:00", "00:30:59"))
            {
                ChkPrintingInk();
            }
            timer.Start();
        }

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
        }

        private void SPnlCountRun()
        {
            try
            {
                SaveFile = Path.GetTempPath() + "SPnlCount.txt";
                //### 2016/07/21 依總經理指示，再將起始日往前推1個月。 ###//
                DateTime startDate = DateTime.Now.AddMonths(-1);
                DataTable result = SPnlCount.SPnlCountRun(startDate);
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
                writerLog.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "   Insert Temp OK!");
                writerLog.Flush();
                SendMail("sm4@ewpcb.com.tw", "每日預估入庫SPnl數量統計", "spnlcount@ewpcb.com.tw",
                    DateTime.Now.ToString("yyyy-MM-dd") + " 每日預估入庫SPnl數量統計！",
                    "每日預估入庫SPnl數量統計，請詳閱附件。" + "<br/>" + "<br/>" +
                    "-----此封郵件由系統所寄出，請勿直接回覆！-----", SaveFile);
            }
            catch (Exception ex)
            {
                writerLog.WriteLine(DateTime.Now.ToString(datetimeFormat) + "  SPnlCountRun()" + ex.Message + "\r\n");
                writerLog.Flush();
            }
        }


        private void EquipMaintainRun()
        {
            try
            {
                SaveFile = Path.GetTempPath() + "EquipMaintain.xls";
                DataTable[] result = new DataTable[] { EquipMaintain.GetWaitMaintain() };
                string[] strSheet = new string[] { "每日待修設備清單" };
                DataTableToExcel(result, strSheet, SaveFile);
                writerLog.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "   Insert Temp OK!");
                writerLog.Flush();
                SendMail("sm4@ewpcb.com.tw", "每日待修設備清單", "equip@ewpcb.com.tw",
                    DateTime.Now.ToString("yyyy-MM-dd") + " 每日待修設備清單！",
                    "每日待修設備清單，請詳閱附件。" + "<br/>" + "<br/>" +
                    "-----此封郵件由系統所寄出，請勿直接回覆！-----", SaveFile);
            }
            catch (Exception ex)
            {
                writerLog.WriteLine(DateTime.Now.ToString(datetimeFormat) + "  EquipMaintainRun()" + ex.Message + "\r\n");
                writerLog.Flush();
            }
        }


        private void ChkDrillHole()
        {
            //設定要扣除的天數
            var decreaseDate = -1;
            //var srcPath = @"\\192.168.1.200\DailyReport4\";
            var srcPath = @"E:\DailyReport4\";
            try
            {
                SaveFile = Path.GetTempPath() + DateTime.Now.AddDays(decreaseDate).ToString("yyyy-MM-dd") +
                    "驗孔數量稽核清單.txt";
                writerResult = new StreamWriter(SaveFile);
                writerResult.WriteLine("批號\t料號\t數量\t驗板數量\t驗板時間(起)\t驗板時間(迄)");
                var ewTB = DFCheckHoleRecord.GetDFRecord(DateTime.Now.AddDays(decreaseDate).ToString("yyyy-MM-dd"));
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
                            writerLog.WriteLine(DateTime.Now.ToString(datetimeFormat) + "  " + ex.Message + "\r" +
                                row["料號"].ToString().Trim() + "+" + chkrow[1].ToString() + "\r");
                            writerLog.Flush();
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
                SendMail("sm4@ewpcb.com.tw", "鑽孔每日驗孔數量稽核清單", "checkhole@ewpcb.com.tw",
                    DateTime.Now.AddDays(decreaseDate).ToString("yyyy-MM-dd") + " 驗孔數量稽核清單！",
                    "鑽孔每日驗孔數量稽核清單，請詳閱附件。" + "<br/>" + "<br/>" +
                    "-----此封郵件由系統所寄出，請勿直接回覆！-----", SaveFile);
            }
            catch (Exception ex)
            {
                writerLog.WriteLine(DateTime.Now.ToString(datetimeFormat) + "  ChkDrillHole()" + ex.Message + "\n\r");
                writerLog.Flush();
            }
        }

        private void ChkPrintingInk()
        {
            SaveFile = Path.GetTempPath() + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + "製令單特殊油墨.txt";
            writerResult = new StreamWriter(SaveFile);
            var result = new DataTable();
            ConnERP ce = new ConnERP(DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd"));
            result = ce.ChkPrintingInk();
            writerResult.WriteLine(string.Format("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}\t{10}\t{11}\t{12}",
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
                result.Columns[11].ColumnName,
                result.Columns[12].ColumnName));
            foreach (DataRow row in result.Rows)
            {
                writerResult.WriteLine(string.Format("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}\t{10}\t{11}\t{12}",
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
                    row[11].ToString().Trim(),
                    row[12].ToString().Trim()));
            }

            //2016-12-07 應工程課長要求，再加入料號需使用到樹脂的檢查
            result.Clear();
            result = ce.ChkIssueResin();
            writerResult.WriteLine();
            writerResult.WriteLine();
            writerResult.WriteLine("==================== 樹脂清單 ====================");
            writerResult.WriteLine();
            writerResult.WriteLine(string.Format("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}",
                result.Columns[0].ColumnName,
                result.Columns[1].ColumnName,
                result.Columns[2].ColumnName,
                result.Columns[3].ColumnName,
                result.Columns[4].ColumnName,
                result.Columns[5].ColumnName,
                result.Columns[6].ColumnName,
                result.Columns[7].ColumnName,
                result.Columns[8].ColumnName,
                result.Columns[9].ColumnName));
            foreach (DataRow row in result.Rows)
            {
                writerResult.WriteLine(string.Format("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}",
                    row[0].ToString().Trim(),
                    row[1].ToString().Trim(),
                    row[2].ToString().Trim(),
                    row[3].ToString().Trim(),
                    row[4].ToString().Trim(),
                    row[5].ToString().Trim(),
                    row[6].ToString().Trim(),
                    row[7].ToString().Trim(),
                    row[8].ToString().Trim(),
                    row[9].ToString().Trim()));
            }
            writerResult.Flush();
            writerResult.Close();
            SendMail("sm4@ewpcb.com.tw", "製令單特殊油墨清單", "chkprintingink@ewpcb.com.tw",
                DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + " 製令單特殊油墨清單！",
                "製令單特殊油墨清單，請詳閱附件。" + "<br/>" + "<br/>" +
                "-----此封郵件由系統所寄出，請勿直接回覆！-----", SaveFile);
        }

        private void GetEveryDayCustomerComplaint()
        {
            try
            {
                SaveFile = Path.GetTempPath() + DateTime.Now.ToString("yyyy-MM-dd") + "品保待處理客訴通知(未逾期).xls";
                ConnEWNAS ce = new ConnEWNAS();
                DataTable[] result = new DataTable[] { ce.EveryDayCustomerComplaint() };
                string[] strSheet = new string[] { "品保待處理客訴通知(未逾期)" };
                DataTableToExcel(result, strSheet, SaveFile);
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
                SendMail("sm4@ewpcb.com.tw", "品保待處理客訴通知(未逾期)", "qagroup@ewpcb.com.tw",
                    DateTime.Now.ToString("yyyy-MM-dd") + " 品保待處理客訴通知(未逾期)",
                    "品保待處理客訴通知(未逾期)，請詳閱附件。" + "<br/>" + "<br/>" +
                    "-----此封郵件由系統所寄出，請勿直接回覆！-----", SaveFile);
            }
            catch (Exception ex)
            {
                writerLog.WriteLine(DateTime.Now.ToString(datetimeFormat) + "  GetEveryDayCustomerComplaint()" + 
                    ex.Message + "\n\r");
                writerLog.Flush();
            }
        }

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
                ConnERP ce = new ConnERP(DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd"));
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
                SaveFile = Path.GetTempPath() + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") +
                    "製令稽核現帳預報廢清單.xls";
                DataTable[] resultTB = new DataTable[] { result };
                string[] strSheet = new string[] { "製令稽核現帳預報廢清單" };
                DataTableToExcel(resultTB, strSheet, SaveFile);
                SendMail("sm4@ewpcb.com.tw", "每日製令稽核現帳預報廢清單", "chkissuescrapwip@ewpcb.com.tw",
                    DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + " 製令稽核現帳預報廢清單",
                    "製令稽核現帳預報廢清單，請詳閱附件。" + "<br/>" + "<br/>" +
                    "-----此封郵件由系統所寄出，請勿直接回覆！-----", SaveFile);
            }
            catch (Exception ex)
            {
                writerLog.WriteLine(DateTime.Now.ToString(datetimeFormat) + "  ChkIssueAndScrapWIP()" + 
                    ex.Message + "\n\r");
                writerLog.Flush();
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
                if (Source[AllSheet].TableName != string.Empty)
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
                writerLog.WriteLine(DateTime.Now.ToString(datetimeFormat) + "  " + ex.Message + "\r\n");
                writerLog.Flush();
            }
            SaveFile.Close();
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
        private void SendMail(string from,string display,string to,string sub,string body,string att)
        {
            //建立寄件者地址與名稱
            MailAddress ReceiverAddress = new MailAddress(from,display);
            //建立收件者地址
            MailAddress SendAddress = new MailAddress(to);
            //建立E-MAIL相關設定與訊息
            MailMessage SendMail = new MailMessage(ReceiverAddress, SendAddress);
            //Mail以HTML格式寄送
            SendMail.IsBodyHtml = true;
            //設定信件內容編碼為UTF8
            SendMail.BodyEncoding = Encoding.UTF8;
            //設定信件主旨編碼為UTF8
            SendMail.SubjectEncoding = Encoding.UTF8;
            //設定信件優先權為普通
            SendMail.Priority = MailPriority.Normal;
            SendMail.Subject = sub;//主旨
            SendMail.Body = body;//內容
            if (att != null)
            {
                //建立附加檔案
                Attachment attachment = new Attachment(att);
                SendMail.Attachments.Add(attachment);//加上附件檔案
            }
            //建立一個信件通訊並設定郵件主機地址與通訊埠號
            SmtpClient MySmtp = new SmtpClient("ms1.ewpcb.com.tw", 25);
            //設定寄件者的帳號與密碼
            MySmtp.Credentials = new NetworkCredential("sm4", "sm4@ew");
            try
            {
                MySmtp.Send(SendMail);
                writerLog.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "   Send Mail OK!");
                writerLog.Flush();
            }
            catch(Exception ex)
            {
                writerLog.WriteLine(DateTime.Now.ToString(datetimeFormat) + "  " + ex.Message + "\r\n");
                writerLog.Flush();
            }
            finally
            {
                MySmtp = null;
                SendMail.Dispose();
                if (att != null)
                {
                    File.Delete(SaveFile);//刪除存放在系統個人暂存區的檔案
                }
            }
        }

        public void Start()
        {
            timer.Enabled = true;
        }

        public void Stop()
        {
            timer.Enabled = false;
            writerLog.Close();
        }
    }
}
