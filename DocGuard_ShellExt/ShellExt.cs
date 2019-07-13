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

/*
 *  Hints    
 *  Don't forget add StrongName to References for build when you develop desktop app!
 *  cd %USERPROFILE%\source\repos\DocGuard\DocGuard_ShellExt\bin\x64\Debug\
 *  Install -   C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm /codebase DocGuard_ShellExt.dll 
 *  Uninstall - C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm /u DocGuard_ShellExt.dll
 */


using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SharpShell.Attributes;
using SharpShell.SharpContextMenu;
using System.IO;
using System.Text.RegularExpressions;
using System.Diagnostics;
using DocGuard_Audit;

namespace DocGuard_ShellExt
{
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.AllFiles)]
    
    public class ShellExt : SharpShell.SharpContextMenu.SharpContextMenu
    {
        public string _fileName { get; private set; }

        protected override bool CanShowMenu()
        {
            return true;
        }

        protected override ContextMenuStrip CreateMenu()
        {
            var menu = new ContextMenuStrip();

            var DocGuard_ShellExtension = new ToolStripMenuItem
            {
                Text = "Analyze with DocGuard",
                Image = Properties.Resource1.DocGuard
            };

            DocGuard_ShellExtension.Click += (sender, args) => Audit();

            //  Add the item to the context menu.
            menu.Items.Add(DocGuard_ShellExtension);

            return menu;
        }

        private void Audit()
        {
            foreach (var filePath in SelectedItemPaths)
            {
                _fileName = filePath;
                string ext = Path.GetExtension(filePath);
                if (Regex.IsMatch(ext, @"\.doc|\.docx|\.xls|\.xlsx", RegexOptions.IgnoreCase))
                {
                    if(DocGuard_Audit.DocGuard.Audit(filePath,"ShellExt"))
                    {
                        string msg = string.Format("Suspicious Attachment: {0}" + Environment.NewLine + Environment.NewLine +
                            "Alert Level: {1}" + Environment.NewLine +
                            "Status: {2}" + Environment.NewLine +
                            "Date: {3}" + Environment.NewLine +
                            "Details: {4}" + Environment.NewLine + Environment.NewLine + "",
                            Environment.NewLine + Path.GetFileName(_fileName), "High", "Logged", DateTime.Now,
                            "Suspicious Module Name =  " + (Infos.randomName ? "Detected" : "Not Detected") + Environment.NewLine +
                            "DDE Attacks =  " + (Infos.ddeString ? "Detected" : "Not Detected") + Environment.NewLine +
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
                            Process.Start(Path.GetDirectoryName(filePath));

                    }
                }
                {
                }
            }
        }


    }
}

