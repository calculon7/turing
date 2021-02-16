using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Turing {
    public class TuringSettings {
        public string ProjectsDirectory { get; set; }
        public string ProjectConfigFileDirectory { get; set; }
        public string RootDownloadDirectory { get; set; }
        public string LogFileDirectory { get; set; }
        public string CoordEmailDirectory { get; set; }
        public string RoamerPath { get; set; }

        public TuringSettings(string turingConfigPath) {
            XDocument doc = XDocument.Load(turingConfigPath);
            XElement root = doc.Root;

            this.ProjectsDirectory = root.Element("ProjectsDirectory").Value;
            this.ProjectConfigFileDirectory = root.Element("ConfigFileDirectory").Value;
            this.RootDownloadDirectory = root.Element("RootDownloadDirectory").Value;
            this.LogFileDirectory = root.Element("LogFileDirectory").Value;
            this.CoordEmailDirectory = root.Element("CoordEmailDirectory").Value;
            this.RoamerPath = root.Element("RoamerPath").Value;
        }
    }

    public class ProjectSettings {
        public enum FileSourceType { 
            Unknown, 
            BIM360Glue, 
            Procore 
        }

        public FileSourceType FileSource { get; set; }
        public string Hostname { get; set; }
        public string ConfigName360Sync { get; set; }
        public string ProjectNumber { get; set; }

        public ProjectSettings(string projectConfigFile) {
            XDocument doc = XDocument.Load(projectConfigFile);
            XElement root = doc.Root;

            this.FileSource = (FileSourceType)Enum.Parse(
                typeof(FileSourceType),
                root.Element("FileSource").Value);

            if (FileSource == FileSourceType.BIM360Glue) {
                this.Hostname = root.Element("Hostname").Value;
            }

            this.ConfigName360Sync = root.Element("ConfigName360Sync").Value;
            this.ProjectNumber = root.Element("ProjectNumber").Value;
        }
    }
}
