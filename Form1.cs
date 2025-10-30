using EpaVM_WF.Properties;
using WinSCP;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EpaVM_WF
{
    public partial class Form1 : Form
    {
        // SFTP 設定 - 從設定檔讀取或使用預設值
        private string Server1 => GetSetting("Server1", "sftp-server1.example.com");
        private string Server2 => GetSetting("Server2", "sftp-server2.example.com");
        private int Port => GetSetting("Port", 22);
        private string UserName => GetSetting("UserName", "your-username");
        private string Password => GetSetting("Password", "your-password");

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                txbRet.Text = "開始時間：" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "\r\n";
                DeleteFile();
                CopyFile();
            }
            catch (Exception ex)
            {
                txbRet.Text += "程式載入錯誤: " + ex.Message + "\r\n";
                WriteLog();
            }
        }

        // 安全的設定讀取方法
        private string GetSetting(string key, string defaultValue)
        {
            try
            {
                var value = Settings.Default[key]?.ToString();
                return string.IsNullOrEmpty(value) ? defaultValue : value;
            }
            catch
            {
                return defaultValue;
            }
        }

        private int GetSetting(string key, int defaultValue)
        {
            try
            {
                if (Settings.Default[key] != null && int.TryParse(Settings.Default[key].ToString(), out int result))
                {
                    return result;
                }
                return defaultValue;
            }
            catch
            {
                return defaultValue;
            }
        }

        private void CopyFile()
        {
            try
            {
                DateTime dtmNow = DateTime.Now.AddDays(-1);
                string targetPath = @"D:\IVRVM\";
                string remotePath = @"/opt/ezvoicetek/ezivr7000/userdata/Epa_Vms/";

                // 確保目標目錄存在
                if (!Directory.Exists(targetPath))
                {
                    Directory.CreateDirectory(targetPath);
                    txbRet.Text += "建立目標目錄: " + targetPath + "\r\n";
                }

                // 從兩個SFTP伺服器下載檔案
                txbRet.Text += "開始從伺服器1下載檔案...\r\n";
                DownloadFromSftpWithWinSCP(Server1, remotePath, targetPath);

                txbRet.Text += "開始從伺服器2下載檔案...\r\n";
                DownloadFromSftpWithWinSCP(Server2, remotePath, targetPath);

                // 檢查是否有下載到檔案
                int fileCount = Directory.GetFiles(targetPath).Length;
                if (fileCount > 0)
                {
                    txbRet.Text += $"成功下載 {fileCount} 個檔案\r\n";

                    // 發送郵件
                    string strSendTo = "johnsonshen@tust.com.tw,linghsu@tust.com.tw,epaduty0800@gmail.com";
                    string subject = "【環保署】語音留言音檔";
                    string body = $"{dtmNow:yyyy-MM-dd} 09:01:00 到 {DateTime.Now:yyyy-MM-dd} 09:00:00 所有留言檔";
                    string cc = "johnsonshen@tust.com.tw,kuanzhou@tust.com.tw,viczhang@tust.com.tw";

                    GmailTo(strSendTo, subject, body, cc);
                }
                else
                {
                    txbRet.Text += "沒有找到符合條件的檔案可下載\r\n";
                }

                // 寄送完刪檔
                CleanupFiles(targetPath);
                WriteLog();
            }
            catch (Exception ex)
            {
                txbRet.Text += "複製檔案錯誤: " + ex.ToString() + "\r\n";
                txbRet.Text += DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "\r\n";
                WriteLog();
            }
        }

        private void DownloadFromSftpWithWinSCP(string server, string remotePath, string targetPath)
        {
            Session session = null;
            try
            {
                // 設定會話選項
                SessionOptions sessionOptions = new SessionOptions
                {
                    Protocol = Protocol.Sftp,
                    HostName = server,
                    UserName = UserName,
                    Password = Password,
                    PortNumber = Port,
                    SshHostKeyPolicy = SshHostKeyPolicy.GiveUpSecurityAndAcceptAny, // 僅用於測試
                    Timeout = TimeSpan.FromSeconds(30)
                };

                session = new Session();

                // 設定會話紀錄
                string logPath = Path.Combine(@"D:\APLOG", $"WinSCP_{server}_{DateTime.Now:yyyyMMddHHmmss}.log");
                session.SessionLogPath = logPath;

                txbRet.Text += $"連接至 SFTP 伺服器: {server}:{Port}\r\n";

                // 連接伺服器
                session.Open(sessionOptions);
                txbRet.Text += "SFTP 連接成功\r\n";

                // 設定傳輸選項
                TransferOptions transferOptions = new TransferOptions
                {
                    TransferMode = TransferMode.Binary,
                    PreserveTimestamp = true
                };

                DateTime cutoffTime = DateTime.Now.AddDays(-1);
                int downloadedFiles = 0;

                // 取得遠端檔案列表
                RemoteDirectoryInfo directory = session.ListDirectory(remotePath);
                txbRet.Text += $"遠端目錄檔案數量: {directory.Files.Count}\r\n";

                foreach (RemoteFileInfo fileInfo in directory.Files)
                {
                    // 跳過目錄和特殊檔案
                    if (fileInfo.IsDirectory || fileInfo.Name == "." || fileInfo.Name == "..")
                        continue;

                    // 檢查檔案名稱長度和時間
                    if (fileInfo.Name.Length >= 14 && fileInfo.LastWriteTime >= cutoffTime)
                    {
                        try
                        {
                            string localFilePath = Path.Combine(targetPath, fileInfo.Name);

                            txbRet.Text += $"下載檔案: {fileInfo.Name} (修改時間: {fileInfo.LastWriteTime:yyyy-MM-dd HH:mm:ss})\r\n";

                            // 下載檔案
                            TransferOperationResult transferResult = session.GetFiles(
                                fileInfo.FullName, localFilePath, false, transferOptions);

                            // 檢查傳輸是否成功
                            if (transferResult.IsSuccess)
                            {
                                downloadedFiles++;
                                txbRet.Text += $"✓ 成功下載: {fileInfo.Name}\r\n";
                            }
                            else
                            {
                                txbRet.Text += $"✗ 下載失敗: {fileInfo.Name}\r\n";
                            }
                        }
                        catch (Exception ex)
                        {
                            txbRet.Text += $"下載檔案 {fileInfo.Name} 時發生錯誤: {ex.Message}\r\n";
                        }
                    }
                }

                txbRet.Text += $"從 {server} 下載完成，共下載 {downloadedFiles} 個檔案\r\n";
            }
            catch (SessionException ex)
            {
                txbRet.Text += $"SFTP 會話錯誤 ({server}): {ex.Message}\r\n";
                if (ex.InnerException != null)
                {
                    txbRet.Text += $"內部錯誤: {ex.InnerException.Message}\r\n";
                }
            }
            catch (Exception ex)
            {
                txbRet.Text += $"SFTP 連線錯誤 ({server}): {ex.Message}\r\n";
            }
            finally
            {
                session?.Dispose();
            }
        }

        // 其餘方法保持不變...
        private void WriteLog()
        {
            // 寫入Log檔
            DateTime dt = DateTime.Now;
            string strDate = dt.ToString("yyyyMMdd");

            string path1 = @"D:\APLOG";
            string logfileName = strDate + "_CopyIVREpaMV_Log.txt";

            if (!System.IO.Directory.Exists(path1))
            {
                System.IO.Directory.CreateDirectory(path1);
            }
            string logfile = System.IO.Path.Combine(path1, logfileName);

            txbRet.Text += "結束時間：" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "\r\n";

            try
            {
                using (StreamWriter sr1 = File.AppendText(logfile))
                {
                    sr1.WriteLine(txbRet.Text.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("寫入Log失敗: " + ex.Message);
            }

            Application.Exit();
        }

        private void DeleteFile()
        {
            string targetPath = @"D:\IVRVM\";

            if (Directory.Exists(targetPath))
            {
                try
                {
                    DirectoryInfo di = new DirectoryInfo(targetPath);
                    FileInfo[] files = di.GetFiles();
                    foreach (FileInfo file in files)
                    {
                        file.Delete();
                    }
                    txbRet.Text += "已刪除舊音檔\r\n";
                }
                catch (Exception ex)
                {
                    txbRet.Text += "刪除舊音檔錯誤: " + ex.Message + "\r\n";
                }
            }
            else
            {
                txbRet.Text += "目標目錄不存在，無需刪除\r\n";
            }
        }

        private void GmailTo(string mailTo, string subject, string body, string strToCC)
        {
            MailMessage msg = null;
            SmtpClient client = null;

            try
            {
                msg = new MailMessage();
                msg.From = new MailAddress("tustedm@gmail.com", "台灣優勢客服", System.Text.Encoding.UTF8);

                // 添加收件人
                foreach (var address in mailTo.Split(','))
                {
                    if (!string.IsNullOrWhiteSpace(address))
                        msg.To.Add(address.Trim());
                }

                // 添加抄送
                if (!string.IsNullOrWhiteSpace(strToCC))
                {
                    foreach (var address in strToCC.Split(','))
                    {
                        if (!string.IsNullOrWhiteSpace(address))
                            msg.CC.Add(address.Trim());
                    }
                }

                msg.Subject = subject;
                msg.SubjectEncoding = System.Text.Encoding.UTF8;
                msg.Body = body;
                msg.BodyEncoding = System.Text.Encoding.UTF8;
                msg.IsBodyHtml = true;

                // 添加附件
                if (Directory.Exists(@"D:\IVRVM\"))
                {
                    string[] files = Directory.GetFiles(@"D:\IVRVM\");
                    if (files.Length > 0)
                    {
                        foreach (var filePath in files)
                        {
                            try
                            {
                                Attachment attachment = new Attachment(filePath);
                                msg.Attachments.Add(attachment);
                                txbRet.Text += $"添加附件: {Path.GetFileName(filePath)}\r\n";
                            }
                            catch (Exception ex)
                            {
                                txbRet.Text += $"添加附件失敗 {Path.GetFileName(filePath)}: {ex.Message}\r\n";
                            }
                        }
                    }
                    else
                    {
                        txbRet.Text += "沒有找到附件檔案\r\n";
                    }
                }

                client = new SmtpClient();
                client.Credentials = new System.Net.NetworkCredential("tustedm@gmail.com", "eeksveuuniuhviwz");
                client.Host = "smtp.gmail.com";
                client.Port = 587;
                client.EnableSsl = true;
                client.Timeout = 30000;

                txbRet.Text += "開始發送郵件...\r\n";
                client.Send(msg);
                txbRet.Text += DateTime.Now + " 郵件寄送成功！\r\n";
            }
            catch (Exception ex)
            {
                txbRet.Text += DateTime.Now + " 郵件寄送失敗: " + ex.Message + "\r\n";
            }
            finally
            {
                msg?.Attachments?.ToList().ForEach(a => a.Dispose());
                msg?.Dispose();
                client?.Dispose();
            }
        }

        private void CleanupFiles(string targetPath)
        {
            try
            {
                if (Directory.Exists(targetPath))
                {
                    Directory.Delete(targetPath, true);
                    txbRet.Text += "已清理暫存檔案和目錄\r\n";
                }
            }
            catch (Exception ex)
            {
                txbRet.Text += $"清理檔案錯誤: {ex.Message}\r\n";
            }
        }
    }
}