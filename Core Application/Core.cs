using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Globalization;
using System.Net;
using System.IO;
using System.Diagnostics;
using System.Configuration;
using System.Threading;

namespace Core_Application
{
    public partial class Core : Form
    {
        public string user = Properties.Settings.Default.User;
        public string password = Properties.Settings.Default.Password;
        public string linkfw = Properties.Settings.Default.LinkFW;
        public string linksoftware = Properties.Settings.Default.LinkSoftware;
        public bool NET = false;
        public bool LOG = false;
        public bool WAS = false;
       
        public Core()
        {
            InitializeComponent();
            notifyIcon1.Visible = true;
            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
            GetFramework("v4.0" , "dotNetFx40.exe");
            DatalogManager("C:/Datalog_Manager", "Datalog_Manager/", "Datalog Delivery");
            IWASManager("NewAutoICTver");
            if (NET == true)
            {
                label1.Text = "Net Framework : OK";
            }
            if (LOG == true)
            {
                label2.Text = "Datalog Manager : OK";
            }
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            DatalogManager("C:/Datalog_Manager", "Datalog_Manager/", "Datalog Delivery");
            IWASManager("NewAutoICTver");
            if (NET == true)
            {
                label1.Text = "Net Framework : OK";
            }
            if (LOG == true)
            {
                label2.Text = "Datalog Manager : OK";
            }
        }
        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            this.WindowState = FormWindowState.Normal;
            notifyIcon1.Visible = true;
        }
        /// <summary>
        /// Auto Download dan silent install
        /// </summary>
        /// <param name="userdni">user ftp</param>
        /// <param name="passworddni">password ftp</param>
        /// <param name="linkfwdni">directori ftp folder penyimpanan Framework</param>
        /// <param name="appnamedni">nama file Framework</param>
        void AutoDnI(string userdni, string passworddni, string linkfwdni, string appnamedni)
        {
            DownloadFile(userdni, passworddni, linkfwdni + appnamedni, AppDomain.CurrentDomain.BaseDirectory + appnamedni);
            SilentInstal(appnamedni);
        }
        
        /// <summary>
        /// Mengecek .Net Framework
        /// </summary>
        /// <param name="version">versi .net framework yang akan dicek</param>
        /// <param name="filename">nama file .net framework</param>
        void GetFramework(string version, string filename)
        {
            try 
            {
                bool NetFramework = false;
                RegistryKey installed_versions = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP");
                string[] version_names = installed_versions.GetSubKeyNames();
                foreach (string ver in version_names)
                {
                    if (ver == version)
                    {
                        NetFramework = true;
                        break;
                    }
                }
                if (NetFramework != true)
                {
                    AutoDnI(user, password, linkfw, filename);
                }
                NET = true;
            }
            catch
            {
                NET = false;
            }
        }
        /// <summary>
        /// Instal Aplikasi yg ada di dalam root applikasi
        /// </summary>
        /// <param name="AppName">nama aplikasi yang akan diinstal</param>
        void SilentInstal(string AppName)
        {
            try
            {
                Process process = new Process();
                process.StartInfo.FileName = AppDomain.CurrentDomain.BaseDirectory + AppName;
                process.StartInfo.Arguments = "/quiet";
                process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                process.Start();
                process.WaitForExit();
                notifyIcon1.BalloonTipText = "Installing Selesai";
                notifyIcon1.ShowBalloonTip(500);
            }
            catch
            {
                notifyIcon1.BalloonTipText = "Installing Gagal";
                notifyIcon1.ShowBalloonTip(500);
            }
        }       
        /// <summary>
        /// Download sebuah file
        /// </summary>
        /// <param name="userName"></param>
        /// <param name="password"></param>
        /// <param name="ftpSourceFilePath">alamat ftp</param>
        /// <param name="localDestinationFilePath">alamat local file</param>
        void DownloadFile(string userName, string password, string ftpSourceFilePath, string localDestinationFilePath)
        {
            try
            {
                int bytesRead = 0;
                byte[] buffer = new byte[2048];
                var requestx = (FtpWebRequest)WebRequest.Create(ftpSourceFilePath);
                requestx.Credentials = new NetworkCredential(userName, password);
                requestx.Method = WebRequestMethods.Ftp.DownloadFile;

                Stream reader = requestx.GetResponse().GetResponseStream();
                FileStream fileStream = new FileStream(localDestinationFilePath, FileMode.Create);
                while (true)
                {
                    bytesRead = reader.Read(buffer, 0, buffer.Length);

                    if (bytesRead == 0)
                    {
                        break;
                    }
                    fileStream.Write(buffer, 0, bytesRead);
                }
                fileStream.Close();
            }
            catch
            {
                notifyIcon1.BalloonTipText = "Gagal Download File Dari FTP Server";
                notifyIcon1.ShowBalloonTip(500);
            }
           
        }
        /// <summary>
        /// Membandingkan isi folder local dengan ftp, jika beda akan diupdate 
        /// </summary>
        /// <param name="localdir">directory local computer yang akan dibandingkan</param>
        /// <param name="ftpdir">directory ftp server</param>
        /// <returns></returns>
        public bool Compare_version(string localdir, string ftpdir)
        {
            try
            {
                List<String> lineslocal = Directory.GetFiles(localdir, "*", SearchOption.AllDirectories).ToList();
                List<string> filelocal = new List<string>();
                foreach (string file in lineslocal)
                {
                    FileInfo mFile = new FileInfo(file);
                    filelocal.Add(mFile.Name);
                }
                int test = filelocal.Count;
                FtpWebRequest listRequest = (FtpWebRequest)WebRequest.Create(ftpdir);
                listRequest.Method = WebRequestMethods.Ftp.ListDirectoryDetails;
                var credentials1 = new NetworkCredential(user, password);
                listRequest.Credentials = credentials1;
                List<string> linesftp = new List<string>();
                //List<string> fileftp = new List<string>();
                using (FtpWebResponse listResponse = (FtpWebResponse)listRequest.GetResponse())
                using (Stream listStream = listResponse.GetResponseStream())
                using (StreamReader listReader = new StreamReader(listStream))
                {
                    while (!listReader.EndOfStream)
                    {
                        linesftp.Add(listReader.ReadLine());
                    }
                }
                int match = 0;
                foreach (string file in linesftp)
                {
                    string[] tokens = file.Split(new[] { ' ' }, 9, StringSplitOptions.RemoveEmptyEntries);
                    string name = tokens[8];
                    for (int a = 0; a <= filelocal.Count; a++)
                    {
                        if (name == filelocal[a])
                        {
                            match++;
                            break;
                        }
                    }
                }
                if (match == filelocal.Count)
                {
                    //semua file sama
                    return true;
                }
                else
                {
                    //ada korup atau beda versi
                    return false;
                }
            }
            catch
            {
                
            }
            return false;
        }
        /// <summary>
        /// Memasukan File name di FTP server kedalam list
        /// </summary>
        /// <param name="url">alamat folder di FTP</param>
        /// <param name="credentials">credential user dan password ftp</param>
        static void ListFtpDirectory(string url, NetworkCredential credentials)
        {
            WebRequest listRequest = WebRequest.Create(url);
            listRequest.Method = WebRequestMethods.Ftp.ListDirectoryDetails;
            listRequest.Credentials = credentials;

            List<string> lines = new List<string>();

            using (WebResponse listResponse = listRequest.GetResponse())
            using (Stream listStream = listResponse.GetResponseStream())
            using (StreamReader listReader = new StreamReader(listStream))
            {
                while (!listReader.EndOfStream)
                {
                    string line = listReader.ReadLine();
                    lines.Add(line);
                }
            }

            foreach (string line in lines)
            {
                string[] tokens =
                    line.Split(new[] { ' ' }, 9, StringSplitOptions.RemoveEmptyEntries);
                string name = tokens[8];
                string permissions = tokens[0];

                if (permissions[0] == 'd')
                {

                    string fileUrl = url + name;
                    ListFtpDirectory(fileUrl + "/", credentials);
                }
                else
                {
                }
            }
        }
        /// <summary>
        /// Menutup Semua Proses
        /// </summary>
        /// <param name="name">nama proses yang ingin ditutup</param>
        void TutupProses(string name)
        {
            foreach (Process clsProcess in Process.GetProcesses())
            {
                if (clsProcess.ProcessName.Contains(name))
                {
                    clsProcess.Kill();
                    clsProcess.WaitForExit();
                    clsProcess.Dispose();
                }
            }
          
        }
        /// <summary>
        /// menghapus semua file yang ada didalam folder
        /// </summary>
        /// <param name="directory">alamat folder yg ingin dihapus</param>
        void DeleteSemua(string directory)
        {
            try
            {
                List<String> local = Directory.GetFiles(directory, "*", SearchOption.AllDirectories).ToList();
                foreach (string file in local)
                {
                    FileInfo mFile = new FileInfo(file);
                    mFile.Delete();
                }
            }
            catch
            {
                notifyIcon1.BalloonTipText = "Gagal Menghapus Directory Local";
                notifyIcon1.ShowBalloonTip(500);
            }
            
        }
        /// <summary>
        /// merunningkan semua file yg ada di directory local
        /// </summary>
        /// <param name="directory">lokasi directory local computer</param>
        void RunningAplikasi(string directory)
        {
            try
            {
                List<String> local = Directory.GetFiles(directory, "*.exe", SearchOption.AllDirectories).ToList();
                foreach (string file in local)
                {
                    FileInfo mFile = new FileInfo(file);
                    Process.Start(mFile.FullName);
                }
            }
            catch
            {
                notifyIcon1.BalloonTipText = "Gagal Running Aplikasi di Directory Local";
                notifyIcon1.ShowBalloonTip(500);
            }
            
        }
        /// <summary>
        /// Auto Update Datalog_Manager Aplikasi
        /// </summary>
        /// <param name="localdir">direktori lokal aplikasi Datalog_Manager</param>
        /// <param name="ftpfolder">Nama Folder ftp tempat menyimpan Datalog_Manager Terupdate</param>
        /// <param name="processname">Nama Process Aplikasi</param>
        void DatalogManager(string localdir, string ftpfolder, string processname)
        {
            bool DM = false;
            if(!File.Exists(localdir))
            {
                System.IO.Directory.CreateDirectory(localdir);
            }
            DM = Compare_version(localdir, linksoftware + ftpfolder);
            if (DM != true)
            {
                TutupProses(processname);
                DeleteSemua(localdir);
                var credentials = new NetworkCredential(user, password);
                DownloadFtpDirectory(linksoftware + ftpfolder, credentials, localdir);
            }
            Process[] pname = Process.GetProcessesByName(processname);
            if (pname.Length == 0)
            {
                RunningAplikasi(localdir);
            }
            if(DM == true)
            {
                LOG = true;
            }
        }
        /// <summary>
        /// Mendownload Seluruh Isi Folder didalam Ftp Server
        /// </summary>
        /// <param name="url">alamat FTP server</param>
        /// <param name="credentials">Credential dari user dan password ftp</param>
        /// <param name="localPath">alamat Hasil download di local Computer</param>
        void DownloadFtpDirectory(string url, NetworkCredential credentials, string localPath)
        {
            try
            {
                FtpWebRequest listRequest = (FtpWebRequest)WebRequest.Create(url);
                listRequest.Method = WebRequestMethods.Ftp.ListDirectoryDetails;
                listRequest.Credentials = credentials;
                List<string> lines = new List<string>();
                using (FtpWebResponse listResponse = (FtpWebResponse)listRequest.GetResponse())
                using (Stream listStream = listResponse.GetResponseStream())
                using (StreamReader listReader = new StreamReader(listStream))
                {
                    while (!listReader.EndOfStream)
                    {
                        lines.Add(listReader.ReadLine());
                    }
                }
                foreach (string line in lines)
                {
                    string[] tokens =
                        line.Split(new[] { ' ' }, 9, StringSplitOptions.RemoveEmptyEntries);
                    string name = tokens[8];
                    string permissions = tokens[0];

                    string localFilePath = Path.Combine(localPath, name);
                    string fileUrl = url + name;

                    if (permissions[0] == 'd')
                    {
                        if (!Directory.Exists(localFilePath))
                        {
                            Directory.CreateDirectory(localFilePath);
                        }

                        DownloadFtpDirectory(fileUrl + "/", credentials, localFilePath);
                    }
                    else
                    {
                        FtpWebRequest downloadRequest = (FtpWebRequest)WebRequest.Create(fileUrl);
                        downloadRequest.Method = WebRequestMethods.Ftp.DownloadFile;
                        downloadRequest.Credentials = credentials;

                        using (FtpWebResponse downloadResponse =
                                  (FtpWebResponse)downloadRequest.GetResponse())
                        using (Stream sourceStream = downloadResponse.GetResponseStream())
                        using (Stream targetStream = File.Create(localFilePath))
                        {
                            byte[] buffer = new byte[10240];
                            int read;
                            while ((read = sourceStream.Read(buffer, 0, buffer.Length)) > 0)
                            {
                                targetStream.Write(buffer, 0, read);
                            }
                        }
                    }
                }
            }
            catch 
            {
                notifyIcon1.BalloonTipText = "Gagal Download Dari FTP Server";
                notifyIcon1.ShowBalloonTip(500);
            }
            
        }
        /// <summary>
        /// Mengecek IWAS version harus diatas 5.5.2 //NewAutoICTver5.5.2
        /// </summary>
        /// <param name="processname">Iwas Process Name</param>
        void IWASManager(string processname)
        {
            bool newiwas = false;
            Process[] processlist = Process.GetProcesses();
            foreach (Process theprocess in processlist)
            {
              if (theprocess.ProcessName.Contains(processname))
              {
                  string prosname = theprocess.ProcessName;
                  string [] version = prosname.Split('r');
                  string[] angka = version[1].Split('.');
                  int major = Convert.ToInt16(angka[0]);
                  int minor = Convert.ToInt16(angka[1]);
                  int revision = Convert.ToInt16(angka[2]);
                  if(major == 5)
                  {
                      if(minor==5)
                      {
                          if (revision>=2)
                          { 
                              newiwas= true;
                          }
                      }
                      if(minor>5)
                  {
                      newiwas = true;
                  }
                  }
                  if(major >5)
                  {
                      newiwas = true;
                  }
              }
            }
            if (newiwas != true)
            {
                TutupProses(processname);
            }
            if (newiwas == true)
            {
                WAS = true;
            }
            if (WAS == true)
            {
                label3.Text = "IWAS - Version : OK";
            }
        }
    }
}