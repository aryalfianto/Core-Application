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
using IWshRuntimeLibrary;

namespace Core_Application
{
    public partial class Core : Form
    {
        public string user = Properties.Settings.Default.User;
        public string password = Properties.Settings.Default.Password;
        public string linksoftware = Properties.Settings.Default.LinkSoftware;
        public string folderftpsoftware = Properties.Settings.Default.FolderFTPSoftware;
        public string ProcessName = Properties.Settings.Default.ProcessName;
        public string LocalDIR = Properties.Settings.Default.Localdir;
       
        public Core()
        {
            InitializeComponent();
            notifyIcon1.Visible = true;
            this.WindowState = FormWindowState.Minimized;
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
            {
                Hide();
                notifyIcon1.Visible = true;
            }
            DatalogManager(LocalDIR, folderftpsoftware, ProcessName);
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            notifyIcon1.Visible = true;
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
                List<String> lineslocal = Directory.GetDirectories(localdir, "*", SearchOption.AllDirectories).ToList();
                List<string> filelocal = new List<string>();
                foreach (string file in lineslocal)
                {
                    FileInfo mFile = new FileInfo(file);
                    filelocal.Add(mFile.Name);
                }   
                FtpWebRequest ftpRequest = (FtpWebRequest)WebRequest.Create(ftpdir);
                ftpRequest.Credentials = new NetworkCredential(user, password);
                ftpRequest.Method = WebRequestMethods.Ftp.ListDirectory;
                FtpWebResponse response = (FtpWebResponse)ftpRequest.GetResponse();
                StreamReader streamReader = new StreamReader(response.GetResponseStream());

                List<string> directories = new List<string>();

                string line = streamReader.ReadLine();
                while (!string.IsNullOrEmpty(line))
                {
                    directories.Add(line);
                    line = streamReader.ReadLine();
                }
                streamReader.Close();
                if (filelocal[0] != directories[0])
                {
                    return false;
                }
                else
                {
                   return true;
                }
            }
            catch
            {
                return true;
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
                var dir = new DirectoryInfo(directory);
                dir.Attributes = dir.Attributes & ~FileAttributes.ReadOnly;
                dir.Delete(true);
            }
            catch
            {
                
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
                List<String> local = Directory.GetFiles(directory, Properties.Settings.Default.FileEXE, SearchOption.AllDirectories).ToList();
                foreach (string file in local)
                {
                    FileInfo mFile = new FileInfo(file);
                    Process.Start(mFile.FullName);
                }
            }
            catch
            {
               
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
            if(!System.IO.File.Exists(localdir))
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
                        using (FtpWebResponse downloadResponse =(FtpWebResponse)downloadRequest.GetResponse())
                        using (Stream sourceStream = downloadResponse.GetResponseStream())
                        using (Stream targetStream = System.IO.File.Create(localFilePath))
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
            }
        }
        void SilentInstal(string AppName)
        {
            try
            {
                Process process = new Process();
                process.StartInfo.FileName = Properties.Settings.Default.LocalVNC + AppName;
                process.StartInfo.Arguments = "/quiet";
                process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                process.Start();
                process.WaitForExit();
                //notifyIcon1.BalloonTipText = "Installing Selesai";
                //notifyIcon1.ShowBalloonTip(500);
            }
            catch
            {
                //notifyIcon1.BalloonTipText = "Installing Gagal";
                //notifyIcon1.ShowBalloonTip(500);
            }
        }

        private void Core_Load(object sender, EventArgs e)
        {
            try
            {
                // Mengecek shortcut sudah dibuat atau belum di dalam folder startup
                if (!System.IO.File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Startup) + "\\Core Apps.lnk"))
                {
                    CreateShortcut(Properties.Settings.Default.Coredir + "/" + "Core Application.exe", Environment.GetFolderPath(Environment.SpecialFolder.Startup) + "\\Core Apps.lnk");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        /// <summary>
        /// Create Windows Shorcut
        /// </summary>
        /// <param name="SourceFile">A file you want to make shortcut to</param>
        /// <param name="ShortcutFile">Path and shorcut file name including file extension (.lnk)</param>
        public static void CreateShortcut(string SourceFile, string ShortcutFile)
        {
            CreateShortcut(SourceFile, ShortcutFile, null, null, null, null);
        }

        /// <summary>
        /// Create Windows Shorcut
        /// </summary>
        /// <param name="SourceFile">A file you want to make shortcut to</param>
        /// <param name="ShortcutFile">Path and shorcut file name including file extension (.lnk)</param>
        /// <param name="Description">Shortcut description</param>
        /// <param name="Arguments">Command line arguments</param>
        /// <param name="HotKey">Shortcut hot key as a string, for example "Ctrl+F"</param>
        /// <param name="WorkingDirectory">"Start in" shorcut parameter</param>
        public static void CreateShortcut(string TargetPath, string ShortcutFile, string Description, string Arguments, string HotKey, string WorkingDirectory)
        {
            // Check necessary parameters first:
            if (String.IsNullOrEmpty(TargetPath))
                throw new ArgumentNullException("TargetPath");
            if (String.IsNullOrEmpty(ShortcutFile))
                throw new ArgumentNullException("ShortcutFile");

            // Create WshShellClass instance:
            var wshShell = new WshShellClass();

            // Create shortcut object:
            IWshRuntimeLibrary.IWshShortcut shorcut = (IWshRuntimeLibrary.IWshShortcut)wshShell.CreateShortcut(ShortcutFile);

            // Assign shortcut properties:
            shorcut.TargetPath = TargetPath;
            shorcut.Description = Description;
            if (!String.IsNullOrEmpty(Arguments))
                shorcut.Arguments = Arguments;
            if (!String.IsNullOrEmpty(HotKey))
                shorcut.Hotkey = HotKey;
            if (!String.IsNullOrEmpty(WorkingDirectory))
                shorcut.WorkingDirectory = WorkingDirectory;

            // Save the shortcut:
            shorcut.Save();
        }
    }
}