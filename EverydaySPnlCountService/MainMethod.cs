using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Data;
using System.Text;
using System.Timers;
using EWPCB_SPnlCountClass;

namespace EverydaySPnlCountService
{
    class MainMethod
    {
        private Timer timer;
        //設定間隔時間為10分鐘
        private double timerInterval = 10 * 60 * 1000;
        private string setTime = "06:30";
        private string datetimeFormat = "yyyy-MM-dd HH:mm:ss";
        private string SaveFile = Path.GetTempPath() + "SPnlCount.txt";
        private StreamWriter writerResult;
        private StreamWriter writerLog;
        
        public MainMethod()
        {
            timer = new Timer();
            timer.Interval = timerInterval;
            timer.AutoReset = false;
            timer.Enabled = false;
            timer.Elapsed += Timer_Elapsed;
            //此方式為每次寫入時，均會直接覆蓋掉原本內容
            writerResult = new StreamWriter(SaveFile);
            //此方式為每次寫入時，持續寫入，不會覆蓋原本內容
            writerLog = File.AppendText("Log.txt");
        }

        private void Timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            timer.Stop();
            if(CheckTime())
            {
                SPnlCountRun();
            }            
            timer.Start();
        }

        private bool CheckTime()
        {
            bool result = false;
            string nowTime = DateTime.Now.ToString("HH:mm");
            if(nowTime==setTime)
            {
                result = true;
                writerLog.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "   (" + nowTime + ")-Return true");
                writerLog.Flush();
            }
            writerLog.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "   (" + nowTime + ")-Return false");
            writerLog.Flush();
            return result;
        }

        private void SPnlCountRun()
        {
            try
            {
                DataTable result = SPnlCount.SPnlCountRun();
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
                SendMail();
            }
            catch (Exception ex)
            {
                writerLog.WriteLine(DateTime.Now.ToString(datetimeFormat) + "  " + ex.Message + "\r\n");
                writerLog.Flush();
            }
        }

        private void SendMail()
        {
            //建立寄件者地址與名稱
            MailAddress ReceiverAddress = new MailAddress("sm4@ewpcb.com.tw", "每日預估入庫SPnl數量統計");
            //建立收件者地址
            MailAddress SendAddress = new MailAddress("spnlcount@ewpcb.com.tw");
            //建立附加檔案
            Attachment attachment = new Attachment(SaveFile);
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
            SendMail.Subject = DateTime.Now.ToString("yyyy-MM-dd") + " 每日預估入庫SPnl數量統計！";//主旨
            SendMail.Body = "每日預估入庫SPnl數量統計，請詳閱附件。" + "<br/>" + "<br/>" +
                "-----此封郵件由系統所寄出，請勿直接回覆！-----";//內容
            SendMail.Attachments.Add(attachment);//加上附件檔案
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
                File.Delete(SaveFile);//刪除存放在系統個人暂存區的檔案
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
