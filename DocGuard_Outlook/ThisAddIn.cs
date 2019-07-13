using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;
using System.Threading;
using System.Text.RegularExpressions;
using DocGuard_Audit;

namespace DocGuard_Outlook
{
    public partial class ThisAddIn
    {
        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;
        Outlook.Items items;
        string attachName = "";

        public static string CreateUniqueTempDirectory()
        {
            var uniqueTempDir = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, Guid.NewGuid().ToString()));
            Directory.CreateDirectory(uniqueTempDir);
            return uniqueTempDir;
        }

        private void LogEvent(string message)
        {
            string eventSource = "DefenseIn DocGuard Outlook";
            //DateTime dt = new DateTime();
            //dt = System.DateTime.UtcNow;
            //message = dt.ToLocalTime() + ": " + message;

            EventLog.WriteEntry(eventSource, message, EventLogEntryType.Warning, 30001);
        }

        void incoming_message(object email)
        {
            Outlook.MailItem mailItem = (Outlook.MailItem)email;

            if (email != null)
            {
                if (mailItem.Attachments.Count > 0)
                {
                    try
                    {
                        for (int i = 1; i <= mailItem.Attachments.Count; i++)
                        {
                            string ext = mailItem.Attachments[i].FileName.Substring(mailItem.Attachments[i].FileName.LastIndexOf('.') + 1);
                            if (Regex.IsMatch(ext, @"doc|docx|xls|xlsx", RegexOptions.IgnoreCase))
                            {
                                attachName = Path.Combine(CreateUniqueTempDirectory(), mailItem.Attachments[i].FileName);
                                mailItem.Attachments[i].SaveAsFile(attachName);
                                try
                                {
                                    try
                                    {
                                        File.Open(attachName, FileMode.Open, FileAccess.Read, FileShare.Read);
                                    }
                                    catch (Exception exi)
                                    {
                                        Thread.Sleep(1500);
                                        LogEvent("First attemtp: " + exi.Message);
                                    }
                                    if (DocGuard_Audit.DocGuard.Audit(attachName,"Outlook"))
                                    {
                                        string msg = string.Format("Suspicious Attachment: {0}" + Environment.NewLine + Environment.NewLine +
                                            "Alert Level: {1}" + Environment.NewLine +
                                            "Status: {2}" + Environment.NewLine +
                                            "Date: {3}" + Environment.NewLine +
                                            "Details: {4}" + Environment.NewLine + "",
                                            Path.GetFileName(attachName), "High", "Logged", DateTime.Now,
                                            "Suspicious Module Name =  " + (DocGuard_Audit.Infos.randomName ? "Detected" : "Not Detected") + Environment.NewLine +
                                                                                "DDE Vulnerability =  " + (Infos.ddeString ? "Detected" : "Not Detected") + Environment.NewLine +
                                                                                "Code Obfuscation =  " + (Infos.obfuscation ? "Detected" : "Not Detected") + Environment.NewLine +
                                                                                "Blacklist Api Usage =  " + (Infos.blaclistApi ? "Detected" : "Not Detected") + Environment.NewLine +
                                                                                "Unviewable Macro Technique =  " + (Infos.unViewable ? "Detected" : "Not Detected") + Environment.NewLine +
                                                                                "Hide Module from VBEditor =  " + (Infos.guiHide ? "Detected" : "Not Detected") + Environment.NewLine +
                                                                                "Macro Files Exported? =  " + (Infos.exportMacro ? "Exported" : "No Export") + Environment.NewLine +
                                                                                Environment.NewLine + Environment.NewLine +
                                                                                "Macro Evasion Technique Detected!" + Environment.NewLine + "Do you want to do examine the findings?");
                                        DialogResult result = MessageBox.Show(msg, "Suspicious Attachment Detected!",
                                            MessageBoxButtons.YesNo,
                                            MessageBoxIcon.Question);
                                        if (result == DialogResult.Yes)
                                            Process.Start(Path.GetDirectoryName(attachName));
                                    }
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.Message);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        string errorInfo = (string)ex.Message.Substring(0, 11);
                        if (errorInfo == "Cannot save")
                        {
                            MessageBox.Show(@"Create Folder C:\TestFileSave");
                        }
                    }
                }
            }
        }


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            outlookNameSpace = this.Application.GetNamespace("MAPI");
            inbox = outlookNameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);

            items = inbox.Items;
            items.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(incoming_message);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
