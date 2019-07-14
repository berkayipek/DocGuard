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
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Interop.Word;
using DocGuard_Audit;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Threading;

namespace DocGuard_Word
{
    public partial class ThisAddIn
    {
        private void Checkout(string fileName)
        {
            string _fileName = fileName;
            try
            {
                try
                {
                    File.Open(_fileName, FileMode.Open, FileAccess.Read, FileShare.Read);
                }
                catch (Exception)
                {
                    Thread.Sleep(1500);
                }
                if (DocGuard_Audit.DocGuard.Audit(_fileName,"Word"))
                {
                    string msg = string.Format("Suspicious Document: {0}" + Environment.NewLine + Environment.NewLine +
                        "Alert Level: {1}" + Environment.NewLine +
                        "Status: {2}" + Environment.NewLine +
                        "Date: {3}" + Environment.NewLine +
                        "Details: {4}" + Environment.NewLine + "",
                        Path.GetFileName(_fileName), "High", "Logged", DateTime.Now,
                        "Suspicious Module Name =  " + (Infos.randomName ? "Detected" : "Not Detected") + Environment.NewLine +
                                                            "DDE Vulnerability =  " + (Infos.ddeString ? "Detected" : "Not Detected") + Environment.NewLine +
                                                            "Code Obfuscation =  " + (Infos.obfuscation ? "Detected" : "Not Detected") + Environment.NewLine +
                                                            "Blacklist Api Usage =  " + (Infos.blaclistApi ? "Detected" : "Not Detected") + Environment.NewLine +
                                                            "Unviewable Macro Technique =  " + (Infos.unViewable ? "Detected" : "Not Detected") + Environment.NewLine +
                                                            "Hide Module from VBEditor =  " + (Infos.guiHide ? "Detected" : "Not Detected") + Environment.NewLine +
                                                            "Macro Files Exported? =  " + (Infos.exportMacro ? "Exported" : "No Export") + Environment.NewLine +
                                                            Environment.NewLine + Environment.NewLine +
                                                            "Suspicious Techniques Detected!" + Environment.NewLine + "Do you want to do examine the findings?");
                    DialogResult result = MessageBox.Show(msg, "Suspicious Document Detected!",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                        Process.Start(Path.GetDirectoryName(fileName));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        void Application_PvWindow(ProtectedViewWindow PvWindow)
        {
            Checkout(PvWindow.Document.FullName);
            
        }

        void Application_DocumentOpen(Word.Document Doc)
        {
            Checkout(Doc.FullName);
        }
    
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.DocumentOpen += new Word.ApplicationEvents4_DocumentOpenEventHandler(Application_DocumentOpen);
            this.Application.ProtectedViewWindowOpen += new Word.ApplicationEvents4_ProtectedViewWindowOpenEventHandler(Application_PvWindow);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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
