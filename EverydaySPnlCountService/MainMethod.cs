using System;
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
            //每日06:30進行SPnl數統計
            if (CheckTime("06:30:00", "06:30:59"))
            {
                SPnlCountRun();
            }
            //每日16:00進行待修設備清單
            else if (CheckTime("16:00:00", "16:00:59"))
            {
                EquipMaintainRun();
            }
            //每日00:10進行鑽孔申報比對驗孔紀錄
            else if (CheckTime("00:10:00", "00:10:59"))
            {
                ChkDrillHole();
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
                writerLog.WriteLine(DateTime.Now.ToString(datetimeFormat) + "  " + ex.Message + "\r\n");
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
                writerLog.WriteLine(DateTime.Now.ToString(datetimeFormat) + "  " + ex.Message + "\r\n");
                writerLog.Flush();
            }
        }


        private void ChkDrillHole()
        {
            try
            {
                SaveFile = Path.GetTempPath() + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + " 驗孔數量稽核清單.txt";
                writerResult = new StreamWriter(SaveFile);
                writerResult.WriteLine("批號\t料號\t數量\t驗板數量\t驗板時間(起)\t驗板時間(迄)\n\r");
                var ewTB = DFCheckHoleRecord.GetDFRecord(DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd"));
                TXTtoTable loadingTxt = new TXTtoTable(@"\\192.168.1.200\DailyReport5\" +
                    DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + ".txt");
                var chkTB = loadingTxt.GetTable();
                foreach (DataRow row in ewTB.Rows)
                {
                    var chkCount = 0;
                    var StartTime = string.Empty;
                    var EndTime = string.Empty;
                    foreach (DataRow chkrow in chkTB.Rows)
                    {
                        if (chkrow["P/N"].ToString() != "")
                        {
                            if (row["料號"].ToString().Trim() == chkrow["P/N"].ToString().ToUpper().Substring(0, 11))
                            {
                                chkCount += Convert.ToInt32(chkrow["BoardCount"]);
                                StartTime = chkrow["StartTime"].ToString();
                                EndTime = chkrow["EndTime"].ToString();
                            }
                        }
                    }
                    if (Convert.ToInt32(row["數量"]) > chkCount)
                    {
                        writerResult.WriteLine(row["批號"].ToString() + "\t" + row["料號"].ToString() + "\t" +
                            row["數量"].ToString() + "\t" + Convert.ToString(chkCount) + "\t" + StartTime + "\t" +
                            EndTime + "\n\r");
                    }
                }
                writerResult.Flush();
                writerResult.Close();
                SendMail("sm4@ewpcb.com.tw", "鑽孔每日驗孔數量稽核清單", "checkhole@ewpcb.com.tw",
                    DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + " 驗孔數量稽核清單！",
                    "鑽孔每日驗孔數量稽核清單，請詳閱附件。" + "<br/>" + "<br/>" +
                    "-----此封郵件由系統所寄出，請勿直接回覆！-----", SaveFile);
            }
            catch (Exception ex)
            {
                writerLog.WriteLine(DateTime.Now.ToString(datetimeFormat) + "  " + ex.Message + "\n\r");
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
