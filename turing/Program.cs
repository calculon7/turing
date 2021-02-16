using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Threading;
using CommandLine;
using Oracle.ManagedDataAccess.Types;

// TODO use file extension db column more

namespace Turing {
    public class Program {
        public static TuringSettings turingSettings;
        public static ProjectSettings projectSettings;
        public static string logFile;
        public static ReaderWriterLock logFileLock = new ReaderWriterLock();
        public static ReaderWriterLock accessDbLock = new ReaderWriterLock();
        public static string turingConfigPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "turing.xml");

        static int Main(string[] args) {
            // load turing settings from xml file in exe location
            turingSettings = new TuringSettings(turingConfigPath);

            return CommandLine.Parser.Default.ParseArguments<CLI.UpdateOptions>(args)
                .MapResult(
                    (CLI.UpdateOptions opts) => Update(opts),
                    errs => 1);
        }

        static int Update(CLI.UpdateOptions opts) {
            // get project number
            if (!Regex.IsMatch(opts.Project, "^\\d{6}$")) {
                throw new Exception("Invalid project number");
            }

            string projectNumber = opts.Project;

            logFile = $"{turingSettings.LogFileDirectory}\\{projectNumber}_{DateTime.Now.ToString("yyyy-MM-dd(HHmm)")}.txt";

            Log("Project number: " + projectNumber);

            // check command line switches
            bool noDl = opts.Nodl;

            if (noDl) {
                Log("360 Sync download disabled");
            }

            bool nwd = opts.Nwd;

            if (!nwd) {
                Log("NWD creation disabled");
            }

            bool showUnknown = opts.ShowUnknown;
            bool showIgnored = opts.ShowIgnored;

            // find turing project config xml file
            var projectConfigFile = Path.Combine(turingSettings.ProjectConfigFileDirectory, projectNumber + ".xml");

            if (!File.Exists(projectConfigFile)) {
                throw new Exception("Project config file does not exist at: " + projectConfigFile);
            }

            var projectDirectory = new DirectoryInfo(turingSettings.ProjectsDirectory).GetDirectories(projectNumber + "*").SingleOrDefault();

            if (projectDirectory == null) {
                throw new Exception("Project path not found");
            }

            // TODO could read from config file
            string projectPath = projectDirectory.FullName;

            // get project info
            Log("Getting project info...");
            var projectInfo = new ProjectInfo(projectPath);

            // find download folder 
            // TODO could read from config file
            var downloadFolderPath = Path.Combine(turingSettings.RootDownloadDirectory, $"{projectInfo.ProjectNumber} - {projectInfo.ProjectName}");
            var downloadFolder = new DirectoryInfo(downloadFolderPath);

            if (!downloadFolder.Exists) {
                Log($"Download folder not found. Creating {downloadFolder.FullName}", MessageType.Warning);
                downloadFolder.Create();
            }

            // get info from config file
            projectSettings = new ProjectSettings(projectConfigFile);

            if (noDl == false) {
                Log("360Sync config file found: " + projectSettings.ConfigName360Sync);

                if (projectSettings.FileSource == ProjectSettings.FileSourceType.BIM360Glue) {
                    // set hostname
                    Log("Setting hostname: " + projectSettings.Hostname);

                    Set360SyncHostname(projectSettings.Hostname);
                }

                // start 360Sync process and wait for it to end
                Log("Running 360Sync...");

                // TODO need to make sure glue source is not disabled
                Run360Sync(projectSettings.ConfigName360Sync);
            }


            // get all files in download folder
            Log("Checking downloaded files...");
            var fileInfos = downloadFolder.GetFiles("*", SearchOption.AllDirectories);

            // create file objects
            var fileObjects = CreateFileObjects(fileInfos, projectInfo);


            // CAD files to process and files to copy
            var newFiles = fileObjects.Where(x => x.TaskType == TaskType.ProcessCad || x.TaskType == TaskType.Copy).ToList();

            // moved to file downloading program
            //foreach (var newFile in newFiles) {
            //    newFile.SetFileStatus(FileStatus.Downloaded);
            //}


            projectInfo.IgnoredFiles = fileObjects.Where(x => x.TaskType == TaskType.Ignore).ToList();

            var cadFilesToProcess = newFiles.OfType<FileObject.CadFile>().ToList();

            Log($"{cadFilesToProcess.Count()} files to update: " + string.Join(", ", cadFilesToProcess.Select(x => x.FileInfo.Name)));


            projectInfo.FilesToCopy = newFiles.OfType<FileObject.NonCadFile>().ToList();

            Log($"{projectInfo.FilesToCopy.Count()} files to copy over: " + string.Join(", ", projectInfo.FilesToCopy.Select(x => x.FileInfo.Name)));


            projectInfo.UnknownFiles = fileObjects.Where(x => x.TaskType == TaskType.Unknown).ToList();

            if (showUnknown) {
                Log($"{projectInfo.UnknownFiles.Count} files not recognized: " + 
                    string.Join(", ", projectInfo.UnknownFiles.Select(x => x.FileInfo.Name)));
            }


            if (showIgnored) {
                Log($"{projectInfo.IgnoredFiles.Count} files ignored: " + string.Join(", ", projectInfo.IgnoredFiles.Select(x => x.FileInfo.Name)));
            }


            if (newFiles.Count == 0) {
                Log("No new files found. Ending process");
                return 0;
            }

            projectInfo.ProcessedFiles = new List<ICadFile>();
            projectInfo.FailedFiles = new List<IFileObject>();

            // execute tasks
            if (cadFilesToProcess.Count + projectInfo.FilesToCopy.Count == 0) {
                Log("No actions to take. Ending process");
                return 0;
            } 
            else {
                if (cadFilesToProcess.Count > 0) {
                    #region process cad files
                    Log("Preparing files");

                    foreach (var cadFileToProcess in cadFilesToProcess) {
                        cadFileToProcess.SetFileStatus(FileStatus.Processing);

                        cadFileToProcess.PrepareForProcessing();
                    }

                    var preparedFiles = cadFilesToProcess.Where(x => x.PrepareSuccessful).ToList();

                    AutoCad.Updating.AssignInboxFolders(preparedFiles);

                    var parallelOptions = new ParallelOptions() { MaxDegreeOfParallelism = 10 };

                    Log("Processing files.");

                    //// for debug 
                    //// process cad files in serial
                    //foreach (var preparedFile in preparedFiles) {
                    //    preparedFile.ExecuteTask();
                    //}

                    // process cad files
                    Parallel.ForEach(preparedFiles, parallelOptions, (preparedFile) => {
                        preparedFile.ExecuteTask();
                    });

                    var processedCadFiles = preparedFiles.Where(x => x.ExecuteSuccessful);

                    // check if cads successful
                    foreach (var cadFileToProcess in cadFilesToProcess) {
                        if (cadFileToProcess.ExecuteSuccessful) {
                            projectInfo.ProcessedFiles.Add(cadFileToProcess);
                            cadFileToProcess.SetFileStatus(FileStatus.Ready);
                        }
                        else {
                            projectInfo.FailedFiles.Add(cadFileToProcess);
                            cadFileToProcess.SetFileStatus(FileStatus.Failed);
                        }
                    }


                    var childFiles = processedCadFiles.Where(x => x.IsChild).ToArray();

                    if (childFiles.Length > 0) {
                        // create parent file objects
                        Log("Creating parent files.");

                        var parents = new List<FileObject.ParentFile>();

                        var childFilesWithDistinctParent = childFiles.DistinctBy(x => x.ParentPath).ToArray();

                        foreach (var childFile in childFilesWithDistinctParent) {
                            var parent = new FileObject.ParentFile() {
                                ProcessedBaseName = Path.GetFileNameWithoutExtension(childFile.ParentPath),
                                InboxSaveAsPath = $"{childFile.ProjectInfo.InboxPath}\\{childFile.ProcessedSubdirectoryName}\\{Path.GetFileNameWithoutExtension(childFile.ParentPath)}.dwg",
                                ProcessedPath = childFile.ParentPath,
                                Level = childFile.Level,
                                Trade = childFile.Trade,
                                ProjectInfo = childFile.ProjectInfo,
                                TemplatePath = childFile.TemplatePath
                            };

                            parents.Add(parent);
                        }


                        foreach (var parent in parents) {
                            parent.SetFileStatus(FileStatus.Processing);

                            var parentProcessedDir = Path.GetDirectoryName(parent.ProcessedPath);

                            // get all parent's children from database Files_IO
                            var childData = Database.Read(
                                "select FILE_INTERNAL_NAME " +
                                "from AIS_DEV.VDC_FILES " +
                                $"where FILE_PROJECT_NUMBER = {projectSettings.ProjectNumber} " +
                                $"and FILE_PARENT_NAME like '{Path.GetFileNameWithoutExtension(parent.ProcessedPath)}'");

                            if (childData.Length == 0) {
                                Log($"No child files found for: {parent.ProcessedBaseName}", MessageType.Error);
                                continue;
                            }

                            OracleString[] allChildrenBaseNames = childData.Select(x => (OracleString)x.Single()).ToArray();

                            // child path are in same directory as parent in vdc-merge directory
                            parent.ChildPaths = allChildrenBaseNames.Select(x => Path.Combine(parentProcessedDir, x.Value + ".dwg")).ToList();

                            parent.ProcessedChildCadFiles = processedCadFiles.Where(x => x.ParentPath == parent.ProcessedPath);
                        }

                        //// debug process parent files
                        //foreach (var parent in parents) {
                        //    AutoCad.Updating.CombineParentFile(parent);
                        //}

                        // process parent files
                        Parallel.ForEach(parents, parallelOptions, parent => {
                            AutoCad.Updating.CombineParentFile(parent);
                        });

                        // check if parents successful
                        foreach (var parent in parents) {
                            if (parent.ExecuteSuccessful) {
                                projectInfo.ProcessedFiles.Add(parent);
                                //SetFileStatus(parent, FileStatus.Ready);
                                parent.SetFileStatus(FileStatus.Ready);
                            }
                            else {
                                projectInfo.FailedFiles.Add(parent);
                                //SetFileStatus(parent, FileStatus.Failed);
                                parent.SetFileStatus(FileStatus.Failed);
                            }
                        }
                    }


                    // TODO this should happen during file processing
                    // if update successful, change Last Processed time of file in database
                    if (projectInfo.ProcessedFiles.Count > 0) {
                        Log("Updating database timestamps");

                        foreach (var dwg in projectInfo.ProcessedFiles) {
                            UpdateLastProcessedTimestamp(dwg.ProcessedPath, projectNumber, useExternalFileName: false);
                        }
                    }

                    if (nwd) {
                        // create nwd for quality checking
                        Log("Creating NWDs");

                        Parallel.ForEach(projectInfo.ProcessedFiles.Select(x => x.ProcessedPath), file => {
                            NwdCreator.Create(file, file.Replace(".dwg", ".nwd"));
                            File.Delete(file.Replace(".dwg", ".nwc"));
                        });
                    }
                    #endregion
                }

                if (projectInfo.FilesToCopy.Count > 0) {
                    #region copy noncad files
                    Log("Copying over files");

                    foreach (var fileObject in projectInfo.FilesToCopy) {
                        var destDir = Path.Combine(fileObject.ProjectInfo.VdcMergePath, "MISC");

                        fileObject.DestinationPath = Path.Combine(destDir, fileObject.FileInfo.Name);

                        Directory.CreateDirectory(destDir);

                        // get last write time of existing file
                        DateTime existingLastWriteTime;

                        if (File.Exists(fileObject.DestinationPath)) {
                            existingLastWriteTime = File.GetLastWriteTime(fileObject.FileInfo.FullName);
                        }
                        else {
                            existingLastWriteTime = DateTime.MinValue;
                        }


                        // copy file to vdc-merge misc folder and overwrite
                        try {
                            File.Copy(fileObject.FileInfo.FullName, fileObject.DestinationPath, true);
                        }
                        catch (IOException){
                            Log($"Could not copy over file {fileObject.FileInfo.Name}", MessageType.Error);
                        }

                        // doesnt update if copying same file, uses same lastWrite timestamp
                        var newLastWriteTime = File.GetLastWriteTime(fileObject.FileInfo.FullName);

                        // check if copy successful
                        if (File.Exists(fileObject.DestinationPath) && DateTime.Compare(existingLastWriteTime, newLastWriteTime) <= 0) {
                            fileObject.CopySuccessful = true;
                            //SetFileStatus(fileObject, FileStatus.Ready);
                            fileObject.SetFileStatus(FileStatus.Ready);
                        }
                        else {
                            projectInfo.FailedFiles.Add(fileObject);
                            //SetFileStatus(fileObject, FileStatus.Failed);
                            fileObject.SetFileStatus(FileStatus.Failed);
                        }
                    }

                    // update database timestamps
                    var successfullyCopiedFiles = projectInfo.FilesToCopy.Where(x => x.CopySuccessful).ToArray();

                    if (successfullyCopiedFiles.Length > 0) {
                        Log("Updating copied files timestamps");

                        foreach (var copiedFile in successfullyCopiedFiles) {
                            UpdateLastProcessedTimestamp(copiedFile.DestinationPath, projectNumber, useExternalFileName: true);
                        }
                    }
                    #endregion
                }

                // create trade file update email
                var attachments = new string[] { };

                var coordEmailfileName = projectNumber + "_" + DateTime.Now.ToString("yyyy-MM-dd(HHmm)") + ".msg";
                var coordEmailFilePath = turingSettings.CoordEmailDirectory + "\\" + coordEmailfileName;

                projectInfo.ProcessedFiles = projectInfo.ProcessedFiles.OrderBy(x => x.FileInfo.Name).ToList();
                projectInfo.FailedFiles = projectInfo.FailedFiles.OrderBy(x => x.FileInfo.Name).ToList();

                if (projectInfo.ProcessedFiles.Count() > 0) {
                    Log("Creating coordination email...");
                    Notification.SaveCoordEmailOutlook(projectInfo, coordEmailFilePath, attachments);
                }

                Log("Sending notification to VDC...");
                Notification.SendTuringNotification(projectInfo, downloadFolder, coordEmailFilePath);
            }

            Log("Process complete!");

            return 0;
        }

        private static IFileObject[] CreateFileObjects(FileInfo[] fileInfos, ProjectInfo projectInfo) {
            var fileData = Database.Read(
                "select FILE_TRADE_ABBREVIATION, FILE_EXTERNAL_NAME, FILE_EXTENSION, FILE_ACTION, FILE_LAST_PROCESSED " +
                "from AIS_DEV.VDC_FILES " +
                $"where FILE_PROJECT_NUMBER = {projectInfo.ProjectNumber}");

            var fileObjects = new IFileObject[fileInfos.Length];

            for (int i = 0; i < fileInfos.Length; i++) {
                var fileInfo = fileInfos[i];

                // get matching database record
                var record = fileData.Where(x => $"{x.ToArray()[1]}.{x.ToArray()[2]}".Equals(fileInfo.Name, StringComparison.OrdinalIgnoreCase)).SingleOrDefault();
                TaskType taskType;

                if (record == null) {
                    taskType = TaskType.Unknown;
                }
                else {
                    var tradeAbbrev = ((OracleString)record[0]).Value;
                    var externalName = ((OracleString)record[1]).Value;
                    var fileExtension = ((OracleString)record[2]).Value;
                    var action = (FileAction)Enum.Parse(typeof(FileAction), ((OracleString)record[3]).Value, true);
                    var lastProcessedDate = (OracleDate)record[4];


                    Trade trade;

                    if (tradeAbbrev.Contains('[') && tradeAbbrev.Contains(']')) {
                        // ignore subtrade
                        // only use string before left bracket
                        // "EL[LITE]" -> "EL"
                        trade = projectInfo.Trades.Where(x => x.Abbreviation == tradeAbbrev.Substring(0, tradeAbbrev.IndexOf('['))).Single();
                    }
                    else {
                        // fallback to blank trade if no match
                        trade = projectInfo.Trades.Where(x => x.Abbreviation == tradeAbbrev).SingleOrDefault() ?? new Trade();
                    }


                    if (action == FileAction.Ignore || trade.IsOwned) {
                        taskType = TaskType.Ignore;
                    }
                    else if (action == FileAction.Clean) {
                        var incomingFileDateModified = fileInfo.LastWriteTime;
                        DateTime lastProcessed;

                        if (lastProcessedDate.IsNull) {
                            lastProcessed = DateTime.MinValue;
                        }
                        else {
                            lastProcessed = (DateTime)lastProcessedDate;
                        }

                        // if incoming file is newer than lastProcessed time
                        if (DateTime.Compare(incomingFileDateModified, lastProcessed) > 0) {
                            taskType = TaskType.ProcessCad;
                        }
                        else {
                            taskType = TaskType.UpToDate;
                        }
                    }
                    else {
                        // file to copy (FileAction.Copy)
                        var incomingFileDateModified = fileInfo.LastWriteTime;
                        DateTime lastProcessed;

                        if (lastProcessedDate.IsNull) {
                            lastProcessed = DateTime.MinValue;
                        }
                        else {
                            lastProcessed = (DateTime)lastProcessedDate;
                        }

                        // if incoming file is newer than lastProcessed time
                        if (DateTime.Compare(incomingFileDateModified, lastProcessed) > 0) {
                            taskType = TaskType.Copy;
                        }
                        else {
                            taskType = TaskType.UpToDate;
                        }
                    }
                }

                IFileObject fileObject = FileObject.Create(fileInfo, taskType, projectInfo);

                fileObjects[i] = fileObject;
            }

            return fileObjects;
        }

        public enum MessageType { 
            Normal, 
            Warning, 
            Error 
        }

        public static void Log(string message, MessageType messageType=MessageType.Normal) {
            logFileLock.AcquireWriterLock(30000);

            try {
                var linePrefix = string.Empty;

                switch (messageType) {
                    case MessageType.Normal:
                        Console.ForegroundColor = ConsoleColor.Gray;
                        break;

                    case MessageType.Warning:
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        linePrefix = "[Warning] ";
                        break;

                    case MessageType.Error:
                        Console.ForegroundColor = ConsoleColor.Red;
                        linePrefix = "[Error] ";
                        break;
                }

                string timestamp = DateTime.Now.ToString("HH:mm:ss");
                string logText = $"{timestamp}: {linePrefix}{message}";

                Console.WriteLine(logText);
                Console.ForegroundColor = ConsoleColor.Gray;

                using (var writer = new StreamWriter(logFile, true)) {
                    writer.WriteLine(logText);
                }
            }
            finally {
                logFileLock.ReleaseWriterLock();
            }
        }

        private static void Run360Sync(string projectConfigName) {
            var processInfo = new ProcessStartInfo {
                FileName = @"filepath",
                Arguments = projectConfigName
            };
            var proc360Sync = Process.Start(processInfo);
            proc360Sync.WaitForExit();
            Log("BIM360Sync process ended with exit code: " + proc360Sync.ExitCode);
        }

        private static void Set360SyncHostname(string hostname) {
            var filePath = @"filepath";

            // XmlDocument won't work because UTF-16 encoding
            string text;
            using (var reader = new StreamReader(filePath)) {
                text = reader.ReadToEnd();
            }

            var newText = Regex.Replace(text, @"<GlueCompanyId>.*</GlueCompanyId>", $"<GlueCompanyId>{hostname}</GlueCompanyId>");
            using (var writer = new StreamWriter(filePath)) {
                writer.Write(newText);
            }
        }

        private static void UpdateLastProcessedTimestamp(string updatedFile, string projectNumber, bool useExternalFileName=false) {
            var basename = Path.GetFileNameWithoutExtension(updatedFile);
            var timestamp = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

            // TODO should probably use FILE_EXTENSION too
            if (useExternalFileName) {
                // for copied files with potentially no internal name
                var write = Database.Write(
                    "update AIS_DEV.VDC_FILES " +
                    $"set FILE_LAST_PROCESSED = TO_DATE('{timestamp}', 'YYYY/MM/DD HH24:MI:SS') " +
                    $"where FILE_EXTERNAL_NAME = '{basename}' " +
                    $"and FILE_PROJECT_NUMBER = {projectNumber}");
                Debug.Assert(write == 1);
            }
            else {
                var write = Database.Write(
                    "update AIS_DEV.VDC_FILES " +
                    $"set FILE_LAST_PROCESSED = TO_DATE('{timestamp}', 'YYYY/MM/DD HH24:MI:SS') " +
                    $"where FILE_INTERNAL_NAME = '{basename}' " +
                    $"and FILE_PROJECT_NUMBER = {projectNumber}");
                Debug.Assert(write == 1);
            }
        }
    }
}
