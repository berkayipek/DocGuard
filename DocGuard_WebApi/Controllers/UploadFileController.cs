using DocGuard_Audit;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;
using System.Xml;

namespace DocGuard_WebApi.Controllers
{
    public class UploadFileController : ApiController
    {
        public string filePath = "";
        string msg = "";

        [HttpPost]
        [Route("api/FileUploading/UploadFile")]
        public async Task<HttpResponseMessage> UploadFile()
        {
            var ctx = HttpContext.Current;
            var root = ctx.Server.MapPath("~/App_Data");
            var provider =
                new MultipartFormDataStreamProvider(root);

            try
            {
                await Request.Content
                    .ReadAsMultipartAsync(provider);

                foreach (var file in provider.FileData)
                {
                    var name = file.Headers
                        .ContentDisposition
                        .FileName;

                    // remove double quotes from string.
                    name = name.Trim('"');

                    var localFileName = file.LocalFileName;
                    filePath = Path.Combine(root, name);

                    File.Move(localFileName, filePath);
                }
            }
            catch (Exception e)
            {
                 
                return new HttpResponseMessage()
                {
                    Content = new StringContent($"Error: {e.Message}", Encoding.UTF8, "application/xml")
                };
            }

            string Extension = Path.GetExtension(filePath);

            if (Regex.IsMatch(Extension, @"\.doc|\.docx|\.xls|\.xlsx", RegexOptions.IgnoreCase))
            {
                try
                {
                    if (DocGuard_Audit.DocGuard.Audit(filePath, "WebApi"))
                    {
                        msg = "<message><msg1> 'Suspicious File: " + filePath + "'</msg1>" +
                            "<msg2> 'Alert Level: Warning'</msg2>" +
                            "<msg3> 'Date: " + DateTime.Now + "'</msg3>" +
                            "<msg4> 'Suspicious Module Name : " + (Infos.randomName ? "Detected" : "Not Detected") + "'</msg4>" +
                            "<msg5> 'DDE Vulnerability :  " + (Infos.ddeString ? "Detected" : "Not Detected") + "'</msg5>" +
                            "<msg6> 'Code Obfuscation :  " + (Infos.obfuscation ? "Detected" : "Not Detected") + "'</msg6>" +
                            "<msg7> 'Blacklist Api Usage :  " + (Infos.blaclistApi ? "Detected" : "Not Detected") + "'</msg7>" +
                            "<msg8> 'Unviewable Macro Technique :  " + (Infos.unViewable ? "Detected" : "Not Detected") + "'</msg8>" +
                            "<msg9> 'Hide Module from VBEditor :  " + (Infos.guiHide ? "Detected" : "Not Detected") + "'</msg9>" +
                            "<msg10> 'Macro Files Exported? :  " + (Infos.exportMacro ? "Exported" : "No Export") + "'</msg10></message>";
                    }
                }
                catch (Exception ex)
                {
                    return new HttpResponseMessage()
                    {
                        Content = new StringContent(ex.Message, Encoding.UTF8, "application/xml")
                    };
                }
            }
            else
            {
                msg = "Unsupported file format!";
                return new HttpResponseMessage()
                {
                    Content = new StringContent(msg, Encoding.UTF8, "application/xml")
                };
            }

            return new HttpResponseMessage()
            {
                Content = new StringContent(msg, Encoding.UTF8, "application/xml")
            };
                    }
    }
}
