// Copyright 2019 DefenseIn Security, Ibrahim Akgul
// ibrahim.akgul@defensein.com - loginit@gmail.com
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <http://www.gnu.org/licenses/>.

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using DocGuard_Audit;

namespace DocGuard_Service
{
    public partial class DocGuardService : ServiceBase
    {
        private FileSystemWatcher docWatcher = null;
        public const string MyServiceName = "DocGuard_Service";
        System.IO.FileStream fs;

        public DocGuardService()
        {
            InitializeComponent();
        }


        private void LogEvent(string message)
        {
            string eventSource = "DefenseIn DocGuard Service ";
            EventLog.WriteEntry(eventSource, message, EventLogEntryType.Warning, 40001);
        }

        private void DocGuard_Scanner(FileSystemEventArgs e)
        {
            try
            {
                // tiny lock control
                try
                {
                    fs = File.Open(e.FullPath, FileMode.Open, FileAccess.Read, FileShare.Read); //fs need for releasing
                }
                catch (Exception exi)
                {
                    Thread.Sleep(1500);
                    LogEvent("First attempt: " + exi.Message);
                }
                if (DocGuard_Audit.DocGuard.Audit(e.FullPath,"Service"))
                {
                    if(fs != null)
                        fs.Close(); // get rid of file

                    string msg = string.Format("Suspicious File: {0}" + Environment.NewLine +
                        "Alert Level: {1}" + Environment.NewLine +
                        "Status: {2}" + Environment.NewLine +
                        "Date: {3}" + Environment.NewLine +
                        "Details: {4}" + Environment.NewLine + "",
                        e.FullPath, "Warning", "Logged", DateTime.Now,
                        "Suspicious Module Name =  " + (Infos.randomName ? "Detected" : "Not Detected") + Environment.NewLine +
                                                            "DDE Vulnerability =  " + (Infos.ddeString ? "Detected" : "Not Detected") + Environment.NewLine +
                                                            "Code Obfuscation =  " + (Infos.obfuscation ? "Detected" : "Not Detected") + Environment.NewLine +
                                                            "Blacklist Api Usage =  " + (Infos.blaclistApi ? "Detected" : "Not Detected") + Environment.NewLine +
                                                            "Unviewable Macro Technique =  " + (Infos.unViewable ? "Detected" : "Not Detected") + Environment.NewLine +
                                                            "Hide Module from VBEditor =  " + (Infos.guiHide ? "Detected" : "Not Detected") + Environment.NewLine +
                                                            "Macro Files Exported? =  " + (Infos.exportMacro ? "Exported" : "No Export") + Environment.NewLine + "",
                                                            "Macro Evasion Technique Detected!");
                    LogEvent(msg);
                }
            }
            catch (Exception ex)
            {
                LogEvent(ex.Message);
            }
        }

        void Checkout(FileSystemEventArgs e)
        {
            string Extension = Path.GetExtension(e.FullPath);

            if (!Path.GetFileNameWithoutExtension(e.FullPath).ToLower().StartsWith("~$"))
            {
                if (Regex.IsMatch(Extension, @"\.doc|\.docx|\.xls|\.xlsx", RegexOptions.IgnoreCase))
                {
                    DocGuard_Scanner(e);
                }
            }
        }

        void Checkout(RenamedEventArgs e)
        {
            string Extension = Path.GetExtension(e.FullPath);

            if (!Path.GetFileNameWithoutExtension(e.FullPath).ToLower().StartsWith("~$"))
            {
                if (Regex.IsMatch(Extension, @"\.doc|\.docx|\.xls|\.xlsx", RegexOptions.IgnoreCase))
                {
                    DocGuard_Scanner(e);
                }

            }
        }

        void OnChanged(object sender, FileSystemEventArgs e)
        {
            Checkout(e);
        }

        void OnRenamed(object sender, RenamedEventArgs e)
        {
            Checkout(e);
        }

        protected override void OnStart(string[] args)
        {
            foreach (var drive in DriveInfo.GetDrives())
            {
                docWatcher = new FileSystemWatcher();
                docWatcher.Path = drive.Name;
                docWatcher.Filter = "*.*";
                docWatcher.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName;
                docWatcher.Created += new FileSystemEventHandler(OnChanged);
                docWatcher.Renamed += new RenamedEventHandler(OnRenamed);

                docWatcher.InternalBufferSize = 8192 * 8; // 64k
                docWatcher.IncludeSubdirectories = true;
                docWatcher.EnableRaisingEvents = true;
            }
        }

        protected override void OnStop()
        {
            // this method is necessary to stop the service while dealing with a heavy file io. 
            // Otherwise we have an unstoppable service ;)
            int iMaxAttempts = 120;
            int iTimeOut = 30000;
            int i = 0;

            while ((!Directory.Exists(docWatcher.Path) || docWatcher.EnableRaisingEvents == false) && i < iMaxAttempts)
            {
                i += 1;
                try
                {
                    docWatcher.EnableRaisingEvents = false;
                    docWatcher.Created -= new FileSystemEventHandler(OnChanged);
                    docWatcher.Renamed -= new RenamedEventHandler(OnRenamed);
                    docWatcher.Dispose();
                    if(!Directory.Exists(docWatcher.Path))
                    {
                        Thread.Sleep(iTimeOut);
                    }
                }
                catch
                {
                    docWatcher.EnableRaisingEvents = false;
                    Thread.Sleep(iTimeOut);
                    LogEvent("Error trying Restart Service");
                }
            }
        }
    }
}
