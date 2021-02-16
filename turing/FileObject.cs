using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Turing {
    public enum FileStatus {
        Unknown,
        Downloaded,
        Processing,
        Ready,
        Failed // add try/finally to make sure status is set to failed if error?
    }

    public interface IFileObject {
        FileInfo FileInfo { get; set; }
        TaskType TaskType { get; set; }
        ProjectInfo ProjectInfo { get; set; }
        void SetFileStatus(FileStatus status);
    }

    public interface ICadFile : IFileObject {
        Trade Trade { get; set; }
        Level Level { get; set; }
        DateTime LastProcessed { get; set; }
        /// <summary>
        /// New timestamped folder that incoming files are placed in per trade. Ex: "HP\\2019-07-30(0856)_HP-01-06"
        /// </summary>
        string ProcessedSubdirectoryName { get; set; }
        /// <summary>
        /// Path that file is saved to after processing but before copying to VDC-merge folder
        /// </summary>
        string InboxSaveAsPath { get; set; }
        /// <summary>
        /// Filepath in VDC-Merge folder
        /// </summary>
        string ProcessedPath { get; set; }
        string TemplatePath { get; set; }
        string ProcessedBaseName { get; set; }
        bool IsChild { get; set; }
        string ParentPath { get; set; }
        void PrepareForProcessing();
        bool PrepareSuccessful { get; set; }
        void ExecuteTask();
        bool ExecuteSuccessful { get; set; }
    }

    public class FileObject : IFileObject {
        public FileInfo FileInfo { get; set; }
        public TaskType TaskType { get; set; }
        public ProjectInfo ProjectInfo { get; set; }
        public FileObject(FileInfo fileInfo, TaskType taskType, ProjectInfo projectInfo) {
            FileInfo = fileInfo;
            TaskType = taskType;
            ProjectInfo = projectInfo;
        }
        public static IFileObject Create(FileInfo fileInfo, TaskType taskType, ProjectInfo projectInfo) {
            switch (taskType) {
                case TaskType.Ignore:
                    return new FileObject(fileInfo, taskType, projectInfo);

                case TaskType.ProcessCad:
                    return new CadFile(fileInfo, taskType, projectInfo);

                case TaskType.UpToDate:
                    return new CadFile(fileInfo, taskType, projectInfo);

                case TaskType.Copy:
                    return new NonCadFile(fileInfo, taskType, projectInfo);

                case TaskType.Unknown:
                default:
                    return new FileObject(fileInfo, taskType, projectInfo);
            }
        }

        public void SetFileStatus(FileStatus status) {
            throw new NotImplementedException();
        }

        public class NonCadFile : IFileObject {
            public FileInfo FileInfo { get; set; }
            public TaskType TaskType { get; set; }
            public ProjectInfo ProjectInfo { get; set; }
            public string DestinationPath { get; set; }
            public bool CopySuccessful { get; set; } = false;
            public NonCadFile(FileInfo fileInfo, TaskType taskType, ProjectInfo projectInfo) {
                FileInfo = fileInfo;
                TaskType = taskType;
                ProjectInfo = projectInfo;
            }

            public void SetFileStatus(FileStatus status) {
                var basename = Path.GetFileNameWithoutExtension(this.FileInfo.FullName);
                var statusString = Enum.GetName(typeof(FileStatus), status);

                Database.Write(
                    "update AIS_DEV.VDC_FILES " +
                    $"set FILE_STATUS = '{statusString}' " +
                    $"where FILE_EXTERNAL_NAME = '{basename}' " +
                    $"and FILE_PROJECT_NUMBER = {this.ProjectInfo.ProjectNumber}");
            }
        }

        public class CadFile : IFileObject, ICadFile {
            public FileInfo FileInfo { get; set; }
            public TaskType TaskType { get; set; }
            public ProjectInfo ProjectInfo { get; set; }
            public Trade Trade { get; set; }
            public Level Level { get; set; }
            public DateTime LastProcessed { get; set; }
            /// <summary>
            /// New timestamped folder that incoming files are placed in per trade. Ex: "HP\\2019-07-30(0856)_HP-01-06"
            /// </summary>
            public string ProcessedSubdirectoryName { get; set; }
            /// <summary>
            /// Filepath in VDC-Merge folder
            /// </summary>
            public string ProcessedPath { get; set; }
            public string TemplatePath { get; set; }
            public bool IsChild { get; set; }
            public string ParentPath { get; set; }
            public bool ExecuteSuccessful { get; set; } = false;
            public bool PrepareSuccessful { get; set; } = false;
            public string InboxSaveAsPath { get; set; }
            public string ProcessedBaseName { get; set; }
            /// <summary>
            /// Filenames of scripts, including file extensions. Ex: ["script1.scr", "script2.scr"]
            /// </summary>
            public List<string> ScriptPaths { get; set; } = new List<string>();

            public string SubTrade { get; set; }

            public CadFile(FileInfo fileInfo, TaskType taskType, ProjectInfo projectInfo) {
                FileInfo = fileInfo;
                TaskType = taskType;
                ProjectInfo = projectInfo;
            }

            public void PrepareForProcessing() {
                AutoCad.Updating.PrepareForProcessing(this);
            }

            public void ExecuteTask() {
                AutoCad.Updating.DoUpdate(this);
            }

            public void SetFileStatus(FileStatus status) {
                var statusString = Enum.GetName(typeof(FileStatus), status);

                var basename = Path.GetFileNameWithoutExtension(this.FileInfo.FullName);

                // TODO user file extension

                if (this.FileInfo.Directory.Parent.Name == "VDC-MERGE") {
                    // file has been processed, FileInfo is now uses internal name
                    var write = Database.Write(
                        "update AIS_DEV.VDC_FILES " +
                        $"set FILE_STATUS = '{statusString}' " +
                        $"where FILE_INTERNAL_NAME = '{basename}' " +
                        $"and FILE_PROJECT_NUMBER = {this.ProjectInfo.ProjectNumber}");
                    Debug.Assert(write == 1);
                }
                else {
                    var write = Database.Write(
                        "update AIS_DEV.VDC_FILES " +
                        $"set FILE_STATUS = '{statusString}' " +
                        $"where FILE_EXTERNAL_NAME = '{basename}' " +
                        $"and FILE_PROJECT_NUMBER = {this.ProjectInfo.ProjectNumber}");
                    Debug.Assert(write == 1);
                }
            }
        }

        public class ParentFile: IFileObject, ICadFile {
            public ProjectInfo ProjectInfo { get; set; }
            public Trade Trade { get; set; }
            public Level Level { get; set; }
            public IEnumerable<ICadFile> ProcessedChildCadFiles { get; set; }
            public List<string> ChildPaths { get; set; }
            public string ProcessedPath { get; set; }
            public string TemplatePath { get; set; }
            public FileInfo FileInfo { get; set; }
            public TaskType TaskType { get; set; }
            public DateTime LastProcessed { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
            public string ProcessedSubdirectoryName { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
            public bool IsChild { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
            public string ParentPath { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
            public void ExecuteTask() {
                throw new NotImplementedException();
            }
            public bool ExecuteSuccessful { get; set; } = false;
            public void PrepareForProcessing() {
                throw new NotImplementedException();
            }

            public void SetFileStatus(FileStatus status) {
                // parent file has no external name or FileInfo to use, use internal name

                var basename = this.ProcessedBaseName;
                var statusString = Enum.GetName(typeof(FileStatus), status);

                Database.Write(
                    "update AIS_DEV.VDC_FILES " +
                    $"set FILE_STATUS = '{statusString}' " +
                    $"where FILE_INTERNAL_NAME = '{basename}' " +
                    $"and FILE_PROJECT_NUMBER = {this.ProjectInfo.ProjectNumber}");
            }

            public bool PrepareSuccessful { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
            public string InboxSaveAsPath { get; set; }
            public string ProcessedBaseName { get; set; }
        }
    }
}
