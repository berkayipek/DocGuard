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
 *       
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenMcdf;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.IO;
using System.IO.Compression;
using Kavod.Vba.Compression;
using System.Runtime.InteropServices;

namespace DocGuard_Audit
{
    public class DocGuard
    {
        #region Globals
        // Name of the generated output file.
        static string outFilename = "";
        static string orgFilename = "";
        static string orgFilePath = "";
        static string _fileName = "";
        static string _sourceApp = "";

        // Compound file that is under editing
        static CompoundFile cf;

        // Byte arrays for holding stream data of file
        static byte[] vbaProjectStream;
        static byte[] dirStream;
        static byte[] projectStream;
        static byte[] wStream;
        #endregion

        #region Audit
        static public bool Audit(string filename,string sourceApp)
        {
            bool lastReport = false;

            // option switches
            bool findHideInGUI = false;
            bool findUnviewableVBA = false;
            bool findRandomName = false;
            bool findBlacklistApiUsage = false;
            bool findObfuscation = false;
            bool findDDEString = true;
            

            bool oldWord = false;
            bool oldExcel = false;

            bool is_OpenXML = false;

            // Temp path to unzip OpenXML files to
            String unzipTempPath = "";

            // OLE Filename (make a copy so we don't overwrite the original)
            outFilename = getOutFilename(filename);
            string oleFilename = outFilename;

            // Attempt to unzip as docm or xlsm OpenXML format
            try
            {
                _fileName = filename;
                _sourceApp = sourceApp;
                orgFilename = getOrgFilename(filename);
                orgFilePath = Path.GetDirectoryName(filename);
                unzipTempPath = CreateUniqueTempDirectory();
                File.Copy(filename, orgFilename);
                ZipFile.ExtractToDirectory(orgFilename, unzipTempPath);
                if (File.Exists(Path.Combine(unzipTempPath, "word", "vbaProject.bin")))
                {
                    oleFilename = Path.Combine(unzipTempPath, "word", "vbaProject.bin");
                }
                else if (File.Exists(Path.Combine(unzipTempPath, "xl", "vbaProject.bin")))
                {
                    oleFilename = Path.Combine(unzipTempPath, "xl", "vbaProject.bin");
                }
                is_OpenXML = true;
            }
            catch (Exception)
            {
                // Not OpenXML format, Maybe 97-2003 format, Make a copy
                if (File.Exists(outFilename)) File.Delete(outFilename);
                if (File.Exists(orgFilename)) File.Delete(orgFilename);

                File.Copy(filename, outFilename);
            }

            // Open OLE compound file for editing
            try
            {
                cf = new CompoundFile(oleFilename, CFSUpdateMode.Update, 0);
            }
            catch (Exception)
            {
                return false;
            }

            // Read relevant streams
            CFStorage commonStorage = cf.RootStorage; // docm or xlsm

            // Office 2003-2007 old format .doc .xls
            if(cf.RootStorage.TryGetStream("workbook") != null)
                wStream = commonStorage.GetStream("workbook").GetData();
            if(cf.RootStorage.TryGetStream("WordDocument") != null)
                wStream = commonStorage.GetStream("WordDocument").GetData();
            if(wStream != null)
            {
                // Read WordDocument/Workbook stream as string
                string wStreamUNICODE = System.Text.Encoding.Unicode.GetString(wStream);
                string wStreamUTF8 = System.Text.Encoding.UTF8.GetString(wStream);
                if (findDDEString)
                {
                    if (blacklistCheck(wStreamUTF8) || blacklistCheck(wStreamUNICODE))
                    {
                        lastReport = true;
                        Infos.ddeString = true;
                    }
                }

            }

            if (cf.RootStorage.TryGetStorage("Macros") != null)
            {
                findHideInGUI = true; findUnviewableVBA = true; findRandomName = true; findBlacklistApiUsage = true; findObfuscation = true;
                oldWord = true;
                commonStorage = cf.RootStorage.GetStorage("Macros"); // .doc
            }
                

            if (cf.RootStorage.TryGetStorage("_VBA_PROJECT_CUR") != null)
            {
                findHideInGUI = true; findUnviewableVBA = true; findRandomName = true; findBlacklistApiUsage = true; findObfuscation = true;
                oldExcel = true;
                commonStorage = cf.RootStorage.GetStorage("_VBA_PROJECT_CUR"); // xls	
            }

            // maybe shoud think for rollback to 070719, remove if
            if(oldWord || oldExcel ||
                cf.RootStorage.TryGetStorage("VBA") != null)
            {
                findHideInGUI = true; findUnviewableVBA = true; findRandomName = true; findBlacklistApiUsage = true; findObfuscation = true;
                vbaProjectStream = commonStorage.GetStorage("VBA").GetStream("_VBA_PROJECT").GetData();
                dirStream = Decompress(commonStorage.GetStorage("VBA").GetStream("dir").GetData());
            }

            if (oldWord || oldExcel ||  cf.RootStorage.TryGetStream("project") != null)
                    projectStream = commonStorage.GetStream("project").GetData();

            if (findHideInGUI || findUnviewableVBA || findRandomName || findBlacklistApiUsage || findObfuscation)
            {
                // Read project stream as string
                string projectStreamString = System.Text.Encoding.UTF8.GetString(projectStream);

                // Find all VBA modules in current file
                List<ModuleInformation> vbaModules = ParseModulesFromDirStream(dirStream);

                List<RootModuleInformation> rootModules = ParseModulesFromRoot(cf);

                if (findBlacklistApiUsage)
                {
                    byte[] streamBytes;
                    string vbaStringCode;
                    foreach (var vbaModule in vbaModules)
                    {
                        streamBytes = commonStorage.GetStorage("VBA").GetStream(vbaModule.orgModuleName).GetData();
                        vbaStringCode = GetVBATextFromModuleStream(streamBytes, vbaModule.textOffset);
                        using (var reader = new StringReader(vbaStringCode))
                        {
                            for (string line = reader.ReadLine(); line != null; line = reader.ReadLine())
                            {
                                if (blacklistCheck(line))
                                {
                                    lastReport = true;
                                    Infos.blaclistApi = true;
                                }
                            }
                        }

                    }
                }


                if (findHideInGUI)
                {
                    foreach (var vbaModule in vbaModules)
                    {
                        if ((vbaModule.moduleName != "ThisDocument") && (vbaModule.moduleName != "ThisWorkbook") && (vbaModule.moduleName != "Sheet1"))
                        {
                            if (!projectStreamString.Contains("Module=" + vbaModule.moduleName))
                            {
                                lastReport = true;
                                Infos.guiHide = true;
                            }
                            else
                            {
                                Infos.guiHide = false;
                            }
                        }
                    }
                }

                if (findUnviewableVBA)
                {
                    if (Regex.IsMatch(projectStreamString, "CMG=\"\"") || Regex.IsMatch(projectStreamString, "GC=\"\""))
                    {
                        lastReport = true;
                        Infos.unViewable = true;
                    }
                    else
                    {
                        Infos.unViewable = false;
                    }

                }

                if (findRandomName)
                {
                    foreach (var vbaModule in vbaModules)
                    {
                        if (vbaModule.orgModuleName != vbaModule.moduleName)
                        {
                            lastReport = true;
                            Infos.randomName = true;
                        }
                        else
                        {
                            Infos.randomName = false;
                        }
                    }
                }

                if (Infos.blaclistApi || Infos.guiHide || Infos.randomName || Infos.unViewable)
                {
                    byte[] streamBytes = { };

                    foreach (var vbaModule in vbaModules)
                    {
                        if (cf.RootStorage.TryGetStorage("Macros") != null)
                            streamBytes = cf.RootStorage.GetStorage("Macros").GetStorage("VBA").GetStream(vbaModule.moduleName).GetData();
                        else if (cf.RootStorage.TryGetStorage("_VBA_PROJECT_CUR") != null)
                            streamBytes = cf.RootStorage.GetStorage("_VBA_PROJECT_CUR").GetStorage("VBA").GetStream(vbaModule.moduleName).GetData(); // xls
                        else
                            streamBytes = commonStorage.GetStorage("VBA").GetStream(vbaModule.moduleName).GetData();

                        if (ExportMacroStrings(vbaModule.moduleName, GetVBATextFromModuleStream(streamBytes, vbaModule.textOffset)))
                        {
                            Infos.exportMacro = true;
                            if (findObfuscation)
                            {
                                double obfuscateRate = ShannonEntropy(GetVBATextFromModuleStream(streamBytes, vbaModule.textOffset));
                                if(obfuscateRate >= 5.3)
                                {
                                    Infos.obfuscation = true;
                                }
                                else
                                {
                                    if(!Infos.obfuscation)
                                        Infos.obfuscation = false;
                                }
                            }
                        }
                        else
                        {
                            Infos.exportMacro = false;
                        }

                    }
                }
            }

            // Commit changes and close file
            cf.Commit();
            cf.Close();

            // Purge unused space in file
            CompoundFile.ShrinkCompoundFile(oleFilename);

            // Zip the file back up as a docm or xlsm
            if (is_OpenXML)
            {
                if (File.Exists(outFilename))
                    File.Delete(outFilename);
                if (File.Exists(orgFilename))
                    File.Delete(orgFilename);
                ZipFile.CreateFromDirectory(unzipTempPath, outFilename);
                // Delete Temporary Files
                Directory.Delete(unzipTempPath, true);
            }
            if (File.Exists(outFilename)) File.Delete(outFilename);
            if (File.Exists(orgFilename)) File.Delete(orgFilename);
            return lastReport;
        }
        #endregion

        #region Helpers
        private static string getOutFilename(String filename)
        {
            string fn = Path.GetFileNameWithoutExtension(filename);
            string ext = ".str";//Path.GetExtension(filename);
            string path = Path.GetDirectoryName(filename);
            return Path.Combine(path, fn + "_Sample" + ext);
        }
        private static string getOrgFilename(String filename)
        {
            string fn = Path.GetFileNameWithoutExtension(filename);
            string ext = ".str";//Path.GetExtension(filename);
            string path = Path.GetDirectoryName(filename);
            return Path.Combine(path, fn + "_Orginal" + ext);
        }

        private static string CreateUniqueTempDirectory()
        {
            var uniqueTempDir = "";
            if (_sourceApp.Contains("ShellExt"))
            {
                uniqueTempDir = Path.GetFullPath(Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString()));
                Directory.CreateDirectory(uniqueTempDir);
                return uniqueTempDir;
            }
            uniqueTempDir = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, Guid.NewGuid().ToString()));
            Directory.CreateDirectory(uniqueTempDir);
            return uniqueTempDir;
        }

        private static bool blacklistCheck(string srcString)
        {
            string[] Blacklist;
            Blacklist = new string[10];

            Blacklist[0] = "OpenProcess";
            Blacklist[1] = "CreateProcess";
            Blacklist[2] = "shell";
            Blacklist[3] = "WriteProcessMemory";
            Blacklist[4] = "CreateRemoteThread";
            Blacklist[5] = "AdjustPrivilege";
            Blacklist[6] = "=MSEXCEL";
            Blacklist[7] = "CreateThread";
            Blacklist[8] = "cmd\\.exe";
            Blacklist[9] = "powershell";

            foreach (string blackList in Blacklist)
            {
                if (Regex.IsMatch(srcString, blackList))
                {
                    return true;
                }
            }
            return false;
        }

        private static string GetVBATextFromModuleStream(byte[] moduleStream, UInt32 textOffset)
        {
            string vbaModuleText = System.Text.Encoding.UTF8.GetString(Decompress(moduleStream.Skip((int)textOffset).ToArray()));

            return vbaModuleText;
        }

        private static bool ExportMacroStrings(string moduleName, string content)
        {
            string directoryName = Path.Combine(orgFilePath,Path.GetFileNameWithoutExtension(_fileName))+ "_artifacts";
            try
            {
                Directory.CreateDirectory(directoryName);
                moduleName = Path.Combine(directoryName, moduleName) + ".bas";
                if (File.Exists(moduleName))
                {
                    File.Delete(moduleName);
                }

                using (FileStream fs = File.Create(moduleName))
                {
                    Byte[] title = new UTF8Encoding(true).GetBytes(content);
                    fs.Write(title, 0, title.Length);
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static List<RootModuleInformation> ParseModulesFromRoot(CompoundFile cf)
        {
            List<RootModuleInformation> modules = new List<RootModuleInformation>();

            RootModuleInformation currentModule = new RootModuleInformation { moduleName = "" };
            Action<CFItem> va = delegate (CFItem item)
            {

                switch (item.Name)
                {
                    case "dir":
                        break;
                    case "_VBA_PROJECT":
                        break;
                    default:
                        currentModule.moduleName = item.Name;
                        modules.Add(currentModule);
                        currentModule = new RootModuleInformation { moduleName = "" };
                        break;
                }
            };
            if (cf.RootStorage.TryGetStorage("Macros") != null)
                cf.RootStorage.GetStorage("Macros").GetStorage("VBA").VisitEntries(va, false);
            else if (cf.RootStorage.TryGetStorage("_VBA_PROJECT_CUR") != null)
                cf.RootStorage.GetStorage("_VBA_PROJECT_CUR").GetStorage("VBA").VisitEntries(va, false); // xls
            else
                cf.RootStorage.GetStorage("VBA").VisitEntries(va, false);

            return modules;
        }

        private static List<ModuleInformation> ParseModulesFromDirStream(byte[] dirStream)
        {
            // 2.3.4.2 dir Stream: Version Independent Project Information
            // https://msdn.microsoft.com/en-us/library/dd906362(v=office.12).aspx
            // Dir stream is ALWAYS in little endian

            List<ModuleInformation> modules = new List<ModuleInformation>();

            int offset = 0;
            UInt16 tag;
            UInt32 wLength;
            ModuleInformation currentModule = new ModuleInformation { moduleName = "", textOffset = 0 };

            while (offset < dirStream.Length)
            {
                tag = GetWord(dirStream, offset);
                wLength = GetDoubleWord(dirStream, offset + 2);

                // The following idiocy is because Microsoft can't stick to their own format specification - taken from Pcodedmp
                if (tag == 9)
                    wLength = 6;
                else if (tag == 3)
                    wLength = 2;

                switch (tag)
                {
                    case 25: // 2.3.4.2.3.2.1 MODULENAME Record
                        currentModule.orgModuleName = System.Text.Encoding.UTF8.GetString(dirStream, (int)offset + 6, (int)wLength);
                        break;
                    case 26: // 2.3.4.2.3.2.3 MODULESTREAMNAME Record
                        currentModule.moduleName = System.Text.Encoding.UTF8.GetString(dirStream, (int)offset + 6, (int)wLength);
                        break;
                    case 49: // 2.3.4.2.3.2.5 MODULEOFFSET Record
                        currentModule.textOffset = GetDoubleWord(dirStream, offset + 6);
                        modules.Add(currentModule);
                        currentModule = new ModuleInformation { moduleName = "", textOffset = 0 };
                        break;
                }

                offset += 6;
                offset += (int)wLength;
            }

            return modules;
        }

        private class ModuleInformation
        {
            public string moduleName; // Name of VBA module stream

            public string orgModuleName;

            public UInt32 textOffset; // Offset of VBA source code in VBA module stream
        }

        private class RootModuleInformation
        {
            public string moduleName;
        }

        private static UInt16 GetWord(byte[] buffer, int offset)
        {
            var rawBytes = new byte[2];
            Array.Copy(buffer, offset, rawBytes, 0, 2);
            return BitConverter.ToUInt16(rawBytes, 0);
        }

        private static UInt32 GetDoubleWord(byte[] buffer, int offset)
        {
            var rawBytes = new byte[4];
            Array.Copy(buffer, offset, rawBytes, 0, 4);
            return BitConverter.ToUInt32(rawBytes, 0);
        }

        private static byte[] Compress(byte[] data)
        {
            var buffer = new DecompressedBuffer(data);
            var container = new CompressedContainer(buffer);
            return container.SerializeData();
        }

        private static byte[] Decompress(byte[] data)
        {
            var container = new CompressedContainer(data);
            var buffer = new DecompressedBuffer(container);
            return buffer.Data;
        }

        //https://codereview.stackexchange.com/a/909
        /// <summary>
        /// returns bits of entropy represented in a given string, per 
        /// http://en.wikipedia.org/wiki/Entropy_(information_theory) 
        /// </summary>
        public static double ShannonEntropy(string s)
        {
            var map = new Dictionary<char, int>();
            foreach (char c in s)
            {
                if (!map.ContainsKey(c))
                    map.Add(c, 1);
                else
                    map[c] += 1;
            }

            double result = 0.0;
            int len = s.Length;
            foreach (var item in map)
            {
                var frequency = (double)item.Value / len;
                result -= frequency * (Math.Log(frequency) / Math.Log(2));
            }
            return result;
        }

        #endregion
    }

}
