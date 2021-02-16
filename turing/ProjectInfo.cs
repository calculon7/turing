using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Oracle.ManagedDataAccess.Types;

namespace Turing {
    public class ProjectInfo {
        public string ProjectPath { get; private set; }
        public string CoordinationPath { get; private set; }
        public string CoordSubmitPath { get; private set; }
        public string InboxPath { get; private set; }
        public string VdcMergePath { get; private set; }

        public int ProjectNumber { get; private set; }
        public string ProjectName { get; private set; }
        public string InternalContact { get; private set; }
        public string ExternalContact { get; private set; }
        public string FtpLocation { get; private set; }

        public string Hostname { get; set; }
        public string Bim360SyncConfigName { get; set; }

        public List<Trade> Trades { get; set; }
        public List<Level> Levels { get; set; }
        public List<FileObject.NonCadFile> FilesToCopy { get; set; }
        public List<IFileObject> IgnoredFiles { get; set; }
        public List<IFileObject> UnknownFiles { get; set; }

        public List<ICadFile> ProcessedFiles { get; set; }
        public List<IFileObject> FailedFiles { get; set; }


        public ProjectInfo(string projectPath) {
            this.ProjectPath = projectPath;

            // project number parsed from project root directory name 
            // TODO use variable from parent function
            string projectDirName = Path.GetFileName(this.ProjectPath);
            Regex regexProjectFolderName = new Regex(@"^(\d{6}) - (.+)$");
            this.ProjectNumber = int.Parse(regexProjectFolderName.Match(projectDirName).Groups[1].Value);

            // can be 2 coordination folders, Q and X
            this.CoordinationPath = Path.Combine(projectPath, @"01 Field Management\04 Coordination");

            // find inbox path
            this.InboxPath = Path.Combine(this.CoordinationPath, "INbox");

            if (!Directory.Exists(InboxPath)) {
                this.InboxPath = Utilities.SwapQXDrive(this.InboxPath);

                if (!Directory.Exists(this.InboxPath)) {
                    throw new Exception("Could not find Inbox path");
                }
            }

            // find coord submit path
            this.CoordSubmitPath = Path.Combine(this.CoordinationPath, "Coord-Submit");

            if (!Directory.Exists(this.CoordSubmitPath)) {
                this.CoordSubmitPath = Utilities.SwapQXDrive(this.CoordSubmitPath);

                if (!Directory.Exists(this.CoordSubmitPath)) {
                    throw new Exception("Could not find Coord-Submit path");
                }
            }

            // create vdc-merge if needed
            this.VdcMergePath = this.InboxPath + "\\VDC-MERGE";
            Directory.CreateDirectory(this.VdcMergePath);


            // get project info from database
            var projectData = Database.Read(
                "select PROJECT_NAME, INTERNAL_EMAIL_GROUP, EXTERNAL_EMAIL_GROUP, PRIMARY_FTP " +
                "from AIS_DEV.VDC_PROJECTS " +
                $"where PROJECT_NUMBER = {this.ProjectNumber}").Single();

            // parse data into ProjectInfo properties
            this.ProjectName =     ((OracleString)projectData[0]).Value;
            this.InternalContact = ((OracleString)projectData[1]).Value;
            this.ExternalContact = ((OracleString)projectData[2]).Value;
            this.FtpLocation =     ((OracleString)projectData[3]).Value;


            // get trades info from database
            var tradesData = Database.Read(
                "select TI_TRADE_ABBREVIATION, TI_OWNED, TI_TRADE_NAME, TI_SUBDIRECTORY, TI_COLOR, TI_XTC, TI_ELEVATION_TO_USE " +
                "from AIS_DEV.VDC_TRADE_INFO " +
                $"where TI_PROJECT_NUMBER = {this.ProjectNumber}");

            this.Trades = new List<Trade>();

            // parse data into Trade properties
            foreach (var record in tradesData) {
                var trade = new Trade {
                    Abbreviation =  ((OracleString)record[0]).Value,
                    IsOwned =       ((OracleDecimal)record[1]).IsZero == false,
                    FullName =      ((OracleString)record[2]).Value,
                    Subdirectory =  ((OracleString)record[3]).Value,
                    Color =         ((OracleString)record[4]).Value,
                    Xtc =           ((OracleDecimal)record[5]).IsZero == false,
                    ElevationToUse = (Trade.ElevationType)Enum.Parse(typeof(Trade.ElevationType), ((OracleString)record[6]).Value, true)
                };

                this.Trades.Add(trade);
            }


            // get level names from database
            var levelsData = Database.Read(
                "select ELEV_LEVEL " +
                "from AIS_DEV.VDC_ELEVATIONS " +
                $"where ELEV_PROJECT_NUMBER = {this.ProjectNumber}");

            this.Levels = new List<Level>();

            foreach (var record in levelsData) {
                var level = new Level {
                    Name = ((OracleString)record[0]).Value
                };

                this.Levels.Add(level);
            }
        }
    }

    public class Trade {
        public enum ElevationType { Unknown, JCC, CDs, RVT }

        public string Abbreviation { get; set; }
        public bool IsOwned { get; set; }
        public string FullName { get; set; }
        public string Subdirectory { get; set; }
        public string Color { get; set; }
        public bool Xtc { get; set; }
        public ElevationType ElevationToUse { get; set; }
    }

    public class Level {
        public string Name { get; set; }
        public string Elevation { get; set; }
        public bool IsRevitElevation { get; set; }
    }

    // for turing object and workflow creation
    public enum TaskType { 
        Unknown, 
        Ignore, 
        Copy, 
        ProcessCad, 
        UpToDate, 
        Post 
    }

    // from database
    // TODO combine enums?
    public enum FileAction {
        Unknown,
        Ignore,
        Copy,
        Clean
    }
}
