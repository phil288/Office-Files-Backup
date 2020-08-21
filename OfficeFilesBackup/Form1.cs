using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.IO;
using System.Management;

namespace OfficeFilesBackup
{
    public partial class Form1 : Form
    {
        Timer timer = new Timer();
        FormWindowState previousState = FormWindowState.Maximized;
        /// <summary>
        /// file modified date
        /// </summary>
        Dictionary<string, string> fileModifiedDate = new Dictionary<string, string>();
        string backupDirectory = Path.Combine(Path.GetPathRoot(Environment.GetFolderPath(Environment.SpecialFolder.System)), "backup");
        public Form1()
        {
            InitializeComponent();
            this.Hide();
            this.ShowInTaskbar = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (!Directory.Exists(backupDirectory)) Directory.CreateDirectory(backupDirectory);
            this.timer.Interval = (60) * 1000;
            this.timer.Tick += timer_Tick;
            this.timer_Tick();
            this.timer.Start();
        }
        void timer_Tick(object sender = null, EventArgs e = null)
        {
            //get all list of all word opened files and try to backup
            Microsoft.Office.Interop.Word.Application objWord;
            Microsoft.Office.Interop.Excel.Application objExcel;
            try
            {
                objWord = (Microsoft.Office.Interop.Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
                for (int i = 0; i < objWord.Windows.Count; i++)
                {
                    object a = i + 1;
                    Window objWin = objWord.Windows.get_Item(ref a);
                    string filename = objWin.Application.Documents[a].FullName;
                    this.backupFile(filename);
                }
            }
            catch (Exception)
            {
                Console.Write("");
            }
            //get a list of all excel opened files and try to backup
            try
            {
                objExcel = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                for (int i = 0; i < objExcel.Windows.Count; i++)
                {
                    object a = i + 1;
                    Microsoft.Office.Interop.Excel.Window objWin = objExcel.Windows.get_Item(a);
                    string filename = objWin.Application.Workbooks[a].FullName;
                    this.backupFile(filename);
                }
            }
            catch (Exception)
            {
                Console.Write("");
            }
            try
            {
                //and now for notepad
                string wmiQuery = string.Format("select CommandLine from Win32_Process where Name='{0}'", "notepad.exe");
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(wmiQuery);
                ManagementObjectCollection retObjectCollection = searcher.Get();
                foreach (ManagementObject retObject in retObjectCollection)
                {
                    string CommandLine = retObject["CommandLine"].ToString();
                    string path = CommandLine.Substring(CommandLine.IndexOf(" ") + 1, CommandLine.Length - CommandLine.IndexOf(" ") - 1);
                    this.backupFile(path);
                }
            }
            catch (Exception)
            {

            }
        }
        /// <summary>
        /// check if the modified date is different than the date that is saved in the filemodified date if it is than create a backup
        /// </summary>
        /// <param name="filename"></param>
        private void backupFile(string filename)
        {
            string modificationDate = "";
            if (File.Exists(filename))
            {
                FileInfo fi = new FileInfo(filename);
                modificationDate = fi.LastWriteTimeUtc.ToString();
                bool backup = false;
                if (!this.fileModifiedDate.ContainsKey(filename))
                {
                    backup = true;
                    this.fileModifiedDate.Add(filename, modificationDate);
                }
                else if (this.fileModifiedDate[filename] != modificationDate)
                {
                    backup = true;
                }
                if (backup)
                {
                    this.fileModifiedDate[filename] = modificationDate;
                    Random random = new Random();//to get a random number
                    //backup the file
                    string backupFilename = Path.Combine(this.backupDirectory,
                        Path.GetFileNameWithoutExtension(filename)
                        + " - " + DateTime.Now.ToString("yyyy-MM-dd hh mm ss tt")
                        + " - " + random.NextDouble()
                        + Path.GetExtension(filename));
                    try
                    {
                        File.Copy(filename, backupFilename, true);
                    }
                    catch (Exception)
                    {
                        Console.Write("");
                    }
                }
            }
        }

        private void Form1_ClientSizeChanged(object sender, EventArgs e)
        {
            if (this.previousState != this.WindowState)
            {
                if (this.WindowState == FormWindowState.Minimized)
                {
                    this.notifyIcon1.Visible = true;
                    this.previousState = this.WindowState;
                    this.ShowInTaskbar = false;
                }
                else
                {
                    this.ShowInTaskbar = true;
                }
            }
            this.previousState = this.WindowState;
        }


        private void notifyIcon1_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {

            }
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            this.notifyIcon1.Visible = false;
            this.Show();
            this.WindowState = FormWindowState.Maximized;
        }
    }
}
