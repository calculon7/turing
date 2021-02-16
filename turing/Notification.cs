using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Turing {
    public class Notification {
        public static readonly string updatingEmailTemplate = @"\\REDACTED\sys\data\everyone\BIM Department\Enterprise\Deployment\Microsoft\Outlook\Templates\VDC_Updating_HardVar.oft";
        public static readonly string postingEmailTemplate = @"\\REDACTED\sys\data\everyone\BIM Department\Enterprise\Deployment\Microsoft\Outlook\Templates\VDC_Posting.oft";

        public static void SendTuringNotification(ProjectInfo projectInfo, DirectoryInfo downloadFolder, string coordEmailFilePath) {

            var copiedFiles = projectInfo.FilesToCopy.Where(x => x.CopySuccessful);

            var subject = $"Turing {Utilities.MorningOrAfternoon()} Notification - {downloadFolder.Name}";

            using (var client = new SmtpClient()) {
                client.Connect("smtp.office365.com", 587, SecureSocketOptions.StartTls);
                client.Authenticate("user@email.com", "password1");

                var message = new MimeMessage();
                message.From.Add(new MailboxAddress("user@email.com"));
                message.To.Add(new MailboxAddress("user@email.com"));
                message.Subject = subject;

                var bodyBuilder = new BodyBuilder();
                bodyBuilder.HtmlBody =
                    "<p><b>Turing Update Notification</b></p>" +
                    $"<p>Project name: {projectInfo.ProjectName}</p>" +
                    $"<p>Project number: {projectInfo.ProjectNumber}</p>" +

                    $"<p><a href=\"{projectInfo.CoordinationPath}\">Coordination folder</a></p>" +
                    $"<p><a href=\"{projectInfo.VdcMergePath}\">Merge folder</a></p>" +

                    $"<p>Items updated successfully: <b>{projectInfo.ProcessedFiles.Count()}</b><br>" +
                    string.Join("<br>", projectInfo.ProcessedFiles.Select(x => $"<a href=\"{x.FileInfo.FullName}\">{x.FileInfo.Name}</a>")) + "</p>" +

                    $"<p>Items failed: <b>{projectInfo.FailedFiles.Count()}</b><br>" +
                    string.Join("<br>", projectInfo.FailedFiles.Select(x => $"<a href=\"{x.FileInfo.FullName}\">{x.FileInfo.Name}</a>")) + "</p>" +

                    $"<p>Items copied: <b>{copiedFiles.Count()}</b><br>" +
                    string.Join("<br>", copiedFiles.Select(x => $"<a href=\"{x.FileInfo.FullName}\">{x.FileInfo.Name}</a>")) + "</p>" +

                    $"<p>Items ignored: <b>{projectInfo.IgnoredFiles.Count()}</b><br>" +
                    string.Join("<br>", projectInfo.IgnoredFiles.Select(x => $"<a href=\"{x.FileInfo.FullName}\">{x.FileInfo.Name}</a>")) + "</p>" +

                    $"<p>Items not recognized: <b>{projectInfo.UnknownFiles.Count()}</b><br>" +
                    string.Join("<br>", projectInfo.UnknownFiles.Select(x => $"<a href=\"{x.FileInfo.FullName}\">{x.FileInfo.Name}</a>")) + "</p>";

                if (File.Exists(coordEmailFilePath)) {
                    bodyBuilder.Attachments.Add(coordEmailFilePath);
                }

                message.Body = bodyBuilder.ToMessageBody();

                client.Send(message);
                client.Disconnect(true);
            }
        }

        public static void SaveCoordEmailOutlook(ProjectInfo projectInfo, string filePath, string[] attachments) {
            Outlook.Application olApp = new Outlook.Application();
            Outlook.MailItem mailItem = olApp.CreateItem(Outlook.OlItemType.olMailItem);

            mailItem.Recipients.Add(projectInfo.InternalContact);
            mailItem.Recipients.ResolveAll();
            mailItem.SentOnBehalfOfName = "user@email.com";
            mailItem.CC = "user@email.com";

            mailItem.Subject = $"{projectInfo.ProjectName} - VDC {Utilities.MorningOrAfternoon()} Update";

            var distinctTrades = projectInfo.ProcessedFiles.Select(x => x.Trade.FullName).Distinct().ToList();
            distinctTrades.Sort();

            // processed and failed cad files
            var processedAndFailedCadFiles = projectInfo.ProcessedFiles.Concat(projectInfo.FailedFiles.OfType<ICadFile>());

            var filesToExpectString = string.Empty;
            foreach (var tradeName in distinctTrades) {
                filesToExpectString += $"{tradeName}:<br>";

                var fileNames = new List<string>();

                // replaces child files with parent and removes duplicates
                foreach (var file in processedAndFailedCadFiles.Where(x => x.Trade.FullName == tradeName)) {
                    if (file is FileObject.CadFile) {
                        if (((FileObject.CadFile)file).IsChild == false) {
                            // if dwg is not a child file
                            fileNames.Add(Path.GetFileName(file.FileInfo.FullName));
                        }
                    } else if (file is FileObject.ParentFile) {
                        fileNames.Add(Path.GetFileName(file.FileInfo.FullName));
                    } else {
                        throw new NotImplementedException();
                    }
                }

                fileNames = fileNames.Distinct().ToList();

                foreach (var fileName in fileNames) {
                    filesToExpectString += $"{fileName}<br>";
                }

                filesToExpectString += "<br>";
            }

            if (!string.IsNullOrWhiteSpace(filesToExpectString)) {
                filesToExpectString = filesToExpectString.Remove(filesToExpectString.LastIndexOf("<br>"));
            }

            //mailItem.HTMLBody =
            //        $"<p style=\"font-family:Arial;font-size:10pt\"><u><i>This is a Distribution Email for:</i></u><br><span style=\"color:rgb(0, 105, 170)\"><b>{projectInfo.ProjectName}</b></span></p>" +
            //        $"<p style=\"font-family:Arial;font-size:10pt\"><u><i>File(s) to expect during this distribution:</i></u><br><span style=\"color:rgb(0, 105, 170)\"><b>{filesToExpectString}</b></span></p>" +
            //        $"<p style=\"font-family:Arial;font-size:10pt\"><u><i>These files were received on:</i></u><br><span style=\"color:rgb(0, 105, 170)\"><b>{DateTime.Now.ToString("MM/dd/yyyy@HHmm")}</b></span></p>" +
            //        $"<p style=\"font-family:Arial;font-size:10pt\"><u><i>Notes/Changes:</i></u><br><span style=\"color:red\">Updated</span></p>" +
            //        $"<br>";

            mailItem.HTMLBody =
                $"<div style=\"font-family:Arial;font-size:10pt\">" +
                $"<i><u>This is a Distribution Email for:</u></i><br>" +
                $"<div style=\"color:rgb(0, 105, 170);font-weight:bold\">{projectInfo.ProjectName}<br>" +
                $"</div><br>" +
                $"<i><u>File(s) to expect during this distribution:</u></i><br>" +
                $"<div style=\"color:rgb(0, 105, 170);font-weight:bold\">{filesToExpectString}<br></div><br>" +
                $"<i><u>These files were last updated on:</u></i><br>" +
                $"<div style=\"color:rgb(0, 105, 170);font-weight:bold\">{DateTime.Now.ToString("MM/dd/yyyy@HHmm")}<br>" +
                $"</div><br>" +
                $"<i><u>Notes/Changes:</u></i><br>" +
                $"<div style=\"color: red\">Updated<br>" +
                $"</div><br>" +
                $"<br>" +
                $"<br>" +
                $"</div>";

            // todo
            // add VDC signature

            foreach (var attachment in attachments) {
                mailItem.Attachments.Add(attachment);
            }

            mailItem.SaveAs(filePath);
        }
    }
}