using Oracle.ManagedDataAccess.Types;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

// TODO PrepareForProcessing() shouldn't query VDC_FILES, VDC_TRADE_INFO, VDC_ELEVATIONS for every file

namespace Turing {

    public class AutoCad {
        public static string acccPath = @"C:\Program Files\Autodesk\AutoCAD 2018\AcCoreConsole.exe";
        public static object zipFileLockObj = new object();

        public static void RunACCC(string cadPath, string scriptPath, bool hidden) {
            string acArgs = $"/i \"{cadPath}\" /s \"{scriptPath}\" /isolate";

            using (Process coreProcess = new Process()) {
                coreProcess.StartInfo.FileName = acccPath;
                coreProcess.StartInfo.Arguments = acArgs;

                // TODO toggle 'hidden' from cli not params
                if (hidden) {
                    coreProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                }

                int c = 0;
                while (!File.Exists(scriptPath)) {
                    Thread.Sleep(1000);
                    c++;

                    if (c == 5) {
                        throw new Exception("Script file not found: " + scriptPath);
                    }
                }

                coreProcess.Start();
                coreProcess.WaitForExit();
            }
        }

        public static class ScriptPaths {
            // TODO make sure all scripts exist

            public static string ScriptsDir = @"\\REDACTED\sys\data\everyone\VDC Group\_BIM Support\Turing\Scripts";

            public static string Updating { get; } = Path.Combine(ScriptsDir, "updating3.scr");
            public static string UpdatingInsert { get; } = Path.Combine(ScriptsDir, "updating_insert.scr");

            public static string A3dChangeColors { get; } = Path.Combine(ScriptsDir, "A3D_CHANGE_COLORS.scr");
            public static string AfChangeColors { get; } = Path.Combine(ScriptsDir, "AF_CHANGE_COLORS.scr");
            public static string ArChangeColors { get; } = Path.Combine(ScriptsDir, "AR_CHANGE_COLORS.scr");

            public static string CombineDwgs { get; } = Path.Combine(ScriptsDir, "combine_dwgs.scr");
        }

        public static class Updating {
            public static void PrepareForProcessing(FileObject.CadFile cadFile) {
                Program.Log($"Preparing {cadFile.FileInfo.Name} for processing.");

                var projectInfo = cadFile.ProjectInfo;
                int projectNumber = projectInfo.ProjectNumber;

                var basename = Path.GetFileNameWithoutExtension(cadFile.FileInfo.FullName);

                // query Files_IO table
                var internalData = Database.Read(
                    "select FILE_LEVEL, FILE_TRADE_ABBREVIATION, FILE_INTERNAL_NAME, FILE_PARENT_NAME, FILE_LAST_PROCESSED, FILE_SCRIPTS " +
                    "from AIS_DEV.VDC_FILES " +
                    $"where FILE_PROJECT_NUMBER = {projectNumber} " +
                    $"and FILE_EXTERNAL_NAME like '{basename}'").SingleOrDefault();

                if (internalData == null) {
                    // no match found
                    Program.Log($"No file match found: {cadFile.FileInfo.Name}", Program.MessageType.Error);
                    //cadFile.SetFileStatus(FileStatus.Failed);
                    return;
                }

                string levelName = ((OracleString)internalData[0]).Value;

                // make own level because may need to use elevation of different level
                cadFile.Level = new Level() {
                    Name = levelName
                };

                var tradeAbbrev = ((OracleString)internalData[1]).Value;

                var tradeMatch = Regex.Match(tradeAbbrev, @"([^[]+)(\[([^]]+)])?");

                var primaryTradeAbbrev = tradeMatch.Groups[1].Value;
                var subtradeAbbrev = string.IsNullOrWhiteSpace(tradeMatch.Groups[3].Value) ? null : tradeMatch.Groups[3].Value; // null if no match/whitespace

                cadFile.Trade = new Trade() {
                    Abbreviation = primaryTradeAbbrev,
                };

                cadFile.SubTrade = subtradeAbbrev;


                cadFile.ProcessedBaseName = ((OracleString)internalData[2]).Value;
                string parentBaseName = null;

                // TODO is whitespace == null?
                if (((OracleString)internalData[3]).IsNull) {
                    cadFile.IsChild = false;
                }
                else {
                    parentBaseName = ((OracleString)internalData[3]).Value;
                    cadFile.IsChild = true;
                }

                if (((OracleDate)internalData[4]).IsNull) {
                    cadFile.LastProcessed = DateTime.MinValue;
                }
                else {
                    cadFile.LastProcessed = ((OracleDate)internalData[4]).Value; // TODO check this
                }


                // check and set autocad scripts to run after processing
                if (((OracleString)internalData[5]).IsNull == false) {
                    var scriptNames = ((OracleString)internalData[5]).Value
                        .Split(',')
                        .Select(scriptName => scriptName.Trim())
                        .Where(scriptName => !string.IsNullOrWhiteSpace(scriptName))
                        .ToArray();

                    var scriptPaths = scriptNames.Select(name => Path.Combine(ScriptPaths.ScriptsDir, name));

                    foreach (var scriptPath in scriptPaths) {
                        if (File.Exists(scriptPath)) {
                            cadFile.ScriptPaths.Add(scriptPath);
                        }
                        else {
                            Program.Log($"Script not found \"{scriptPath}\"", Program.MessageType.Error);
                            //cadFile.SetFileStatus(FileStatus.Failed);
                            return;
                        }
                    }
                }


                // TODO instead of building level for each cadfile, build all levels first then assign as necessary (if possible with using level above, etc)
                // query Trade_info table
                var tradeData = Database.Read(
                    "select TI_TRADE_NAME, TI_SUBDIRECTORY, TI_COLOR, TI_XTC, TI_ELEVATION_TO_USE " +
                    "from AIS_DEV.VDC_TRADE_INFO " +
                    $"where TI_PROJECT_NUMBER = {projectNumber} " +
                    $"and TI_TRADE_ABBREVIATION like '{cadFile.Trade.Abbreviation}'").SingleOrDefault();

                if (tradeData == null) {
                    // no match found
                    Program.Log($"Trade {cadFile.Trade.Abbreviation} not found for {projectNumber}", Program.MessageType.Error);
                    //cadFile.SetFileStatus(FileStatus.Failed);
                    return;
                }

                cadFile.Trade.FullName = ((OracleString)tradeData[0]).Value;
                cadFile.Trade.Subdirectory = ((OracleString)tradeData[1]).Value;
                cadFile.Trade.Color = ((OracleString)tradeData[2]).Value;
                cadFile.Trade.Xtc = ((OracleDecimal)tradeData[3]).IsZero == false;

                switch (((OracleString)tradeData[4]).Value.ToUpper()) {
                    case "JCC":
                        cadFile.Trade.ElevationToUse = Trade.ElevationType.JCC;
                        break;

                    case "CDS":
                        cadFile.Trade.ElevationToUse = Trade.ElevationType.CDs;
                        break;

                    case "RVT":
                        cadFile.Trade.ElevationToUse = Trade.ElevationType.RVT;
                        break;

                    default:
                        // invalid elevation type
                        Program.Log($"Invalid elevation type: {((OracleString)tradeData[4]).Value}", Program.MessageType.Error);
                        return;
                }

                // vdc-merge trade subdirectory
                var processedDir = $"{projectInfo.VdcMergePath}\\{cadFile.Trade.FullName}";
                Directory.CreateDirectory(processedDir);

                cadFile.ProcessedPath = $"{processedDir}\\{cadFile.ProcessedBaseName}.dwg";

                if (cadFile.IsChild) {
                    cadFile.ParentPath = $"{processedDir}\\{parentBaseName}.dwg"; // TODO rename to parent processed path?
                }


                // use elevation of level below for structural files
                var strucTradeAbbrevs = new string[] { "STC", "STC(2D)", "STE", "STE(2D)" };

                cadFile.Level.IsRevitElevation = cadFile.Trade.ElevationToUse == Trade.ElevationType.RVT;
                bool isStructuralFile = strucTradeAbbrevs.Any(x => x.Equals(cadFile.Trade.Abbreviation, StringComparison.OrdinalIgnoreCase));

                string query = $"select {(cadFile.Level.IsRevitElevation ? "ELEV_REVIT_ELEVATION" : "ELEV_ELEVATION")} " +
                               $"from AIS_DEV.VDC_ELEVATIONS " +
                               $"where ELEV_PROJECT_NUMBER = {projectNumber} " +
                               $"and {(isStructuralFile ? "ELEV_LEVEL_ABOVE" : "ELEV_LEVEL")} = '{cadFile.Level.Name}'";

                var elevData = Database.Read(query).SingleOrDefault().SingleOrDefault();

                if (elevData == null) {
                    // no match found
                    Program.Log($"Elevation not found for {cadFile.ProcessedBaseName}", Program.MessageType.Error);
                    //cadFile.SetFileStatus(FileStatus.Failed);
                    return;
                }

                cadFile.Level.Elevation = ((OracleString)elevData).Value;

                cadFile.PrepareSuccessful = true;
            }

            public static void AssignInboxFolders(IEnumerable<FileObject.CadFile> files) {
                var tradeAbbrevs = files.Select(x => x.Trade.Abbreviation).Distinct().ToList();

                var timeStamp = DateTime.Now.ToString("yyyy-MM-dd(HHmm)");

                var archTradeAbbrevs = new string[] { "AF", "AR", "A3D" };
                var strucTradeAbbrevs = new string[] { "STC", "STC(2D)", "STE", "STE(2D)" };

                foreach (var tradeAbbrev in tradeAbbrevs) {
                    // groups AF,AR,A3D and STC,STC(2D),STE,STE(2D) in inbox
                    string inboxTradeAbbrev;

                    if (archTradeAbbrevs.Any(x => x.Equals(tradeAbbrev, StringComparison.OrdinalIgnoreCase))) {
                        inboxTradeAbbrev = "ARCH";
                    }
                    else if (strucTradeAbbrevs.Any(x => x.Equals(tradeAbbrev, StringComparison.OrdinalIgnoreCase))) {
                        inboxTradeAbbrev = "STRC";
                    }
                    else {
                        inboxTradeAbbrev = tradeAbbrev;
                    }

                    var filesOfTrade = files.Where(x => x.Trade.Abbreviation == tradeAbbrev);
                    var levelNamesOfTrade = filesOfTrade.Select(x => x.Level.Name).Distinct().ToList();
                    levelNamesOfTrade.Sort();

                    string levelsString;

                    if (levelNamesOfTrade.Count() < 5) {
                        levelsString = string.Join(",", levelNamesOfTrade);
                    }
                    else {
                        levelsString = $"{levelNamesOfTrade.First()}-{levelNamesOfTrade.Last()}";
                    }

                    var dirName = $"{timeStamp}_{tradeAbbrev}-{levelsString}";

                    foreach (var file in filesOfTrade) {
                        file.ProcessedSubdirectoryName = $"{inboxTradeAbbrev}\\{dirName}";
                        file.InboxSaveAsPath = $"{file.ProjectInfo.InboxPath}\\{file.ProcessedSubdirectoryName}\\{file.ProcessedBaseName}.dwg";
                    }
                }
            }

            public static void DoUpdate(FileObject.CadFile cadFile) {
                var projectInfo = cadFile.ProjectInfo;

                // create new inbox folder
                var newInboxSubdirectoryInfo = Directory.CreateDirectory($"{projectInfo.InboxPath}\\{cadFile.ProcessedSubdirectoryName}");
                var backupDirectoryInfo = Directory.CreateDirectory(Path.Combine(newInboxSubdirectoryInfo.FullName, "Backup"));
                var scriptsDirectoryInfo = Directory.CreateDirectory(Path.Combine(backupDirectoryInfo.FullName, "Scripts"));
                var refTagsDirectoryInfo = Directory.CreateDirectory(Path.Combine(backupDirectoryInfo.FullName, "Reftags"));

                var inboxFilePath = $"{newInboxSubdirectoryInfo.FullName}\\{cadFile.FileInfo.Name}";

                // copy file to inbox folder
                File.Copy(cadFile.FileInfo.FullName, inboxFilePath, true);
                cadFile.FileInfo = new FileInfo(inboxFilePath);

                lock (zipFileLockObj) {
                    // create or add to backup zip
                    var zipPath = Path.Combine(backupDirectoryInfo.FullName, "backup.zip");

                    if (File.Exists(zipPath)) {
                        // add file to zip
                        while (Utilities.IsFileReady(zipPath) == false) {
                            //Program.Log(Program.MessageType.Warning, $"Waiting for file {zipPath}.");
                            Thread.Sleep(1000);
                        }

                        // todo // sometimes this still fails, lock and IsFileReady don't always work
                        try {
                            using (var zipArchive = ZipFile.Open(zipPath, ZipArchiveMode.Update)) {
                                zipArchive.CreateEntryFromFile(cadFile.FileInfo.FullName, cadFile.FileInfo.Name);
                            }
                        } 
                        catch (IOException) {
                            Program.Log($"Could not backup file {cadFile.FileInfo.Name}, trying again", Program.MessageType.Warning);
                            try {
                                // try again
                                do {
                                    Thread.Sleep(1000);
                                } while (Utilities.IsFileReady(zipPath) == false);

                                using (var zipArchive = ZipFile.Open(zipPath, ZipArchiveMode.Update)) {
                                    zipArchive.CreateEntryFromFile(cadFile.FileInfo.FullName, cadFile.FileInfo.Name);
                                }
                            } 
                            catch (IOException) {
                                Program.Log($"Could not backup file {cadFile.FileInfo.Name}", Program.MessageType.Error);
                            }
                        }
                    } 
                    else {
                        // create zip and add file
                        using (var zipToOpen = new FileStream(zipPath, FileMode.Create)) {
                            using (var zipArchive = new ZipArchive(zipToOpen, ZipArchiveMode.Update)) {
                                zipArchive.CreateEntryFromFile(cadFile.FileInfo.FullName, cadFile.FileInfo.Name);
                            }
                        }
                    }
                }

                // start cad processing
                try {
                    // make refTag file
                    string refTagScriptPath;
                    string templatePath;

                    if (string.IsNullOrWhiteSpace(cadFile.SubTrade)) {
                        // if no subtrade
                        refTagScriptPath = $"{scriptsDirectoryInfo.FullName}\\{cadFile.Trade.Abbreviation}-Lev{cadFile.Level.Name}_createRefTag.scr";
                        var refTagType = GetRefTagType(cadFile);
                        templatePath = CreateRefTag(refTagType, cadFile.Level.Name, refTagScriptPath,
                            $"{refTagsDirectoryInfo.FullName}\\{cadFile.Trade.Abbreviation}-Lev{cadFile.Level.Name}_reftag.dwg").FullName;
                    } 
                    else {
                        refTagScriptPath = $"{scriptsDirectoryInfo.FullName}\\{cadFile.Trade.Abbreviation}[{cadFile.SubTrade}]-Lev{cadFile.Level.Name}_createRefTag.scr";
                        var refTagType = GetRefTagType(cadFile);
                        templatePath = CreateRefTag(refTagType, cadFile.Level.Name, refTagScriptPath,
                            $"{refTagsDirectoryInfo.FullName}\\{cadFile.Trade.Abbreviation}[{cadFile.SubTrade}]-Lev{cadFile.Level.Name}_reftag.dwg").FullName;
                    }

                    cadFile.TemplatePath = templatePath;

                    CleanIncoming(cadFile);
                    InsertCad(cadFile);
                    RunPerFileScripts(cadFile);
                }
                catch (Exception ex) {
                    Program.Log(ex.Message, Program.MessageType.Error);
                }
            }

            private static void RunPerFileScripts(FileObject.CadFile cadFile) {
                foreach (var script in cadFile.ScriptPaths) {
                    var scriptPath = Path.Combine(ScriptPaths.ScriptsDir, script);

                    RunACCC(cadFile.FileInfo.FullName, scriptPath, false);
                }
            }

            private static void CleanIncoming(FileObject.CadFile cadFile) {
                #region edit script text
                string fullCleanIncomingText;

                using (StreamReader reader = new StreamReader(ScriptPaths.Updating)) {
                    fullCleanIncomingText = reader.ReadToEnd();
                }

                // set color
                if (cadFile.Trade.Color.Equals("nochange", StringComparison.OrdinalIgnoreCase)) {
                    fullCleanIncomingText = fullCleanIncomingText.Replace("%CHANGECOLOR%", string.Empty);
                }
                else {
                    fullCleanIncomingText = fullCleanIncomingText.Replace("%CHANGECOLOR%", $"color {cadFile.Trade.Color} * ");
                }

                #region architectural scripts
                switch (cadFile.Trade.Abbreviation) {
                    case "A3D":
                        string a3dScriptText;

                        using (StreamReader reader = new StreamReader(ScriptPaths.A3dChangeColors)) {
                            a3dScriptText = reader.ReadToEnd();
                        }

                        fullCleanIncomingText = fullCleanIncomingText.Replace(";%ARCH_SCRIPT%", a3dScriptText);
                        break;

                    case "AF":
                        string afScriptText;

                        using (StreamReader reader = new StreamReader(ScriptPaths.AfChangeColors)) {
                            afScriptText = reader.ReadToEnd();
                        }

                        fullCleanIncomingText = fullCleanIncomingText.Replace(";%ARCH_SCRIPT%", afScriptText);
                        break;

                    case "AR":
                        string arScriptText;

                        using (StreamReader reader = new StreamReader(ScriptPaths.ArChangeColors)) {
                            arScriptText = reader.ReadToEnd();
                        }

                        fullCleanIncomingText = fullCleanIncomingText.Replace(";%ARCH_SCRIPT%", arScriptText);
                        break;

                    default:
                        break;
                }
                #endregion

                #region elevate and blockall
                if ((cadFile.Trade.ElevationToUse == Trade.ElevationType.CDs || cadFile.Trade.ElevationToUse == Trade.ElevationType.RVT) && cadFile.Level.Elevation != "0'0") {
                    string blockName = $"{cadFile.Level.Name}-Insert-{ DateTime.Now.Millisecond}";
                    string blockallString = $"_.-block {blockName} 0,0,0 ALL \n_.-insert {blockName} 0,0,0 1 1 0";

                    fullCleanIncomingText = fullCleanIncomingText.Replace(";block", blockallString);

                    string formattedElev;
                    if (!cadFile.Level.Elevation.StartsWith("-")) {
                        // if positive, make negative
                        formattedElev = cadFile.Level.Elevation.Insert(0, "-"); // 0,0,-0'0
                    } else {
                        // if negative, make positive
                        formattedElev = cadFile.Level.Elevation.TrimStart('-');
                    }

                    string moveString = $"_.MOVE ALL  0,0,0 0,0,{formattedElev}";

                    fullCleanIncomingText = fullCleanIncomingText.Replace(";elevate", moveString);

                }
                #endregion

                #region xtc
                string acadPath = null;
                if (cadFile.Trade.Xtc) {
                    string xtcString = string.Format("-view sw regenall -exporttoautocad p .\ns -ACAD\n\n");
                    fullCleanIncomingText = fullCleanIncomingText.Replace(";-exporttoautocad", xtcString);

                    acadPath = Regex.Replace(cadFile.FileInfo.FullName, "\\.dwg", "-ACAD.dwg", RegexOptions.IgnoreCase);

                    // if -ACAD.dwg already exists
                    if (File.Exists(acadPath)) {
                        try {
                            // delete existing -ACAD file for overwriting
                            File.Delete(acadPath);
                        }
                        catch {
                            Program.Log($"The existing file {cadFile.FileInfo.FullName.Replace(".dwg", " - ACAD.dwg")} " +
                                $"could not be deleted. Make sure the file is not in use and delete before proceeding.", Program.MessageType.Error);
                        }
                    }
                }
                else {
                    fullCleanIncomingText += "qsave\n";
                }
                #endregion

                // write final cleanincoming script file
                //var finalScriptPath = $"{cadFile.FileInfo.Directory.FullName}\\Backup\\Scripts\\{cadFile.Trade.Abbreviation}[{cadFile.SubTrade}]-Lev{cadFile.Level.Name}.scr";
                string finalScriptPath;

                if (string.IsNullOrWhiteSpace(cadFile.SubTrade)) {
                    // if no subtrade
                    finalScriptPath = $"{cadFile.FileInfo.Directory.FullName}\\Backup\\Scripts\\{cadFile.Trade.Abbreviation}-Lev{cadFile.Level.Name}_clean.scr";
                }
                else {
                    finalScriptPath = $"{cadFile.FileInfo.Directory.FullName}\\Backup\\Scripts\\{cadFile.Trade.Abbreviation}[{cadFile.SubTrade}]-Lev{cadFile.Level.Name}_clean.scr";
                }

                File.WriteAllText(finalScriptPath, fullCleanIncomingText);
                #endregion

                // run accoreconsole with new script text
                RunACCC(cadFile.FileInfo.FullName, finalScriptPath, true);

                // use acad file from now on
                if (cadFile.Trade.Xtc) {
                    if (File.Exists(acadPath)) {
                        cadFile.FileInfo = new FileInfo(acadPath);
                    }
                    else {
                        throw new Exception("-ACAD file was not exported sucessfully.");
                    }
                }
            }

            private static void InsertCad(FileObject.CadFile cadFile) {
                string insertScriptText;

                using (StreamReader reader = new StreamReader(ScriptPaths.UpdatingInsert)) {
                    insertScriptText = reader.ReadToEnd();
                }

                insertScriptText = insertScriptText.Replace("%insertblock%", "INSERT_" + cadFile.Trade.Abbreviation);
                insertScriptText = insertScriptText.Replace("%FILENAME%", $"\"{cadFile.FileInfo.FullName}\"");
                insertScriptText = insertScriptText.Replace("%REVIEWFILEPATH%", $"\"{cadFile.InboxSaveAsPath}\"");

                // write final cleanincoming script file
                //string finalScriptPath = Path.Combine(cadFile.FileInfo.Directory.FullName, "Backup\\Scripts", cadFile.Trade.Abbreviation + "-Lev" + cadFile.Level.Name + "_insert.scr");
                string finalScriptPath;

                if (string.IsNullOrWhiteSpace(cadFile.SubTrade)) {
                    // if no subtrade
                    finalScriptPath = $"{cadFile.FileInfo.Directory.FullName}\\Backup\\Scripts\\{cadFile.Trade.Abbreviation}-Lev{cadFile.Level.Name}_insert.scr";
                }
                else {
                    finalScriptPath = $"{cadFile.FileInfo.Directory.FullName}\\Backup\\Scripts\\{cadFile.Trade.Abbreviation}[{cadFile.SubTrade}]-Lev{cadFile.Level.Name}_insert.scr";
                }

                File.WriteAllText(finalScriptPath, insertScriptText);

                if (File.Exists(cadFile.InboxSaveAsPath)) {
                    File.Delete(cadFile.InboxSaveAsPath);
                }

                Program.Log($"Inserting {cadFile.FileInfo.Name} into {Path.GetFileName(cadFile.TemplatePath)} and saving to {Path.GetFileName(cadFile.InboxSaveAsPath)}.");
                var timeBeforeStart = DateTime.Now;

                // run insert script on template and save to vdc-merge subdirectory
                RunACCC(cadFile.TemplatePath, finalScriptPath, true);

                if (File.Exists(cadFile.InboxSaveAsPath) && DateTime.Compare(new FileInfo(cadFile.InboxSaveAsPath).LastWriteTime, timeBeforeStart) > 0) {
                    // if file exists and has been modified recently
                    cadFile.FileInfo = new FileInfo(cadFile.InboxSaveAsPath);
                }
                else {
                    throw new Exception($"File {cadFile.FileInfo.Name} could not be inserted into template.");
                }

                // copy to Vdc-Merge
                File.Copy(cadFile.InboxSaveAsPath, cadFile.ProcessedPath, true);

                if (File.Exists(cadFile.ProcessedPath) && DateTime.Compare(new FileInfo(cadFile.ProcessedPath).LastWriteTime, timeBeforeStart) > 0) {
                    // if file exists and has been modified recently
                    cadFile.FileInfo = new FileInfo(cadFile.ProcessedPath);
                    cadFile.ExecuteSuccessful = true;
                }
                else {
                    throw new Exception($"File {cadFile.FileInfo.Name} could not be copied to Vdc-Merge.");
                }
            }

            public static void CombineParentFile(FileObject.ParentFile parent) {
                var projectInfo = parent.ProjectInfo;

                // do not process if any child files were unsuccessful
                if (!parent.ProcessedChildCadFiles.All(x => x.ExecuteSuccessful)) {
                    Program.Log($"Parent file {parent.ProcessedPath} was not processed because one or more of its children were not processed successfully.", Program.MessageType.Error);
                    return;
                }

                #region edit script text
                // make combine script
                //string combineScriptPath = projectInfo.InboxPath + "\\" +
                //                           parent.ChildCadFilesProcessed.First().InboxSubdirectoryName + "\\Backup\\Scripts\\" +
                //                           Path.GetFileNameWithoutExtension(parent.ProcessedPath) + "_combine.scr";

                var combineScriptPath = $"{projectInfo.InboxPath}\\{parent.ProcessedChildCadFiles.First().ProcessedSubdirectoryName}" +
                                    $"\\Backup\\Scripts\\{parent.Trade.Abbreviation}_{parent.Level.Name}_combine.scr";

                // insert all children that exist into template on own layer
                string combineScriptText;

                using (StreamReader reader = new StreamReader(ScriptPaths.CombineDwgs)) {
                    combineScriptText = reader.ReadToEnd();
                }

                var layersAndInsertsText = string.Empty;

                // leave out child files that do not exist
                foreach (string childPath in parent.ChildPaths.Where(x => File.Exists(x))) {
                    var layerName = Regex.Match(Path.GetFileNameWithoutExtension(childPath), @"\d{6}_(.+)-Lev").Groups[1].Value;

                    if (string.IsNullOrWhiteSpace(layerName)) {
                        Program.Log($"Could not parse subtrade for {childPath}, falling back to basename.", Program.MessageType.Error);
                        layerName = Path.GetFileNameWithoutExtension(childPath);
                    }

                    layersAndInsertsText += $"-LAYER MAKE \"{layerName}\"\n\n";
                    layersAndInsertsText += $"-insert \"{childPath}\"\n0,0,0\n\n\n\n";
                }

                layersAndInsertsText = layersAndInsertsText.TrimEnd('\n');

                combineScriptText = combineScriptText.Replace("%LAYERSANDINSERTS%", layersAndInsertsText);
                // save to inbox folder
                combineScriptText = combineScriptText.Replace("%COMBINEDFILEPATH%", "\"" + parent.InboxSaveAsPath + "\"");

                File.WriteAllText(combineScriptPath, combineScriptText);
                #endregion

                if (File.Exists(parent.InboxSaveAsPath)) {
                    File.Delete(parent.InboxSaveAsPath);
                }

                Program.Log("Combining files " + string.Join(", ", parent.ChildPaths.Where(x => File.Exists(x)).Select(x => Path.GetFileName(x))) + " into: " + Path.GetFileName(parent.ProcessedPath));
                
                var timeBeforeStart = DateTime.Now;

                RunACCC(parent.TemplatePath, combineScriptPath, true);

                if (File.Exists(parent.InboxSaveAsPath) &&
                    DateTime.Compare(new FileInfo(parent.InboxSaveAsPath).LastWriteTime, timeBeforeStart) > 0) {
                    // if parent file exists and has been modified recently
                    parent.FileInfo = new FileInfo(parent.InboxSaveAsPath);
                }
                else {
                    Program.Log("Parent file could not be created.", Program.MessageType.Error);
                }

                // copy to Vdc-Merge
                File.Copy(parent.InboxSaveAsPath, parent.ProcessedPath, true);

                if (File.Exists(parent.ProcessedPath) && DateTime.Compare(new FileInfo(parent.ProcessedPath).LastWriteTime, timeBeforeStart) > 0) {
                    // if parent file exists and has been modified recently
                    parent.FileInfo = new FileInfo(parent.ProcessedPath);
                    parent.ExecuteSuccessful = true;
                }
                else {
                    Program.Log("Parent file could not be copied to Vdc-Merge folder.", Program.MessageType.Error);
                }
            }

            public static RefTagType GetRefTagType(ICadFile cadFile) {
                switch (cadFile.Trade.Abbreviation.ToUpper()) {
                    case "A3D":
                        return RefTagType.A3D;

                    case "AF":
                        return RefTagType.AF;

                    case "AR":
                        return RefTagType.AR;

                    case "EL":
                    case "ES":
                        return RefTagType.EL;

                    case "FP":
                    case "FS":
                        return RefTagType.FP;

                    case "HD":
                    case "HO":
                        return RefTagType.HD;

                    case "HP":
                    case "HS":
                        return RefTagType.HP;

                    case "PL":
                    case "PS":
                        return RefTagType.PL;

                    case "STC":
                    case "STC(2D)":
                        return RefTagType.STC;

                    case "STE":
                    case "STE(2D)":
                        return RefTagType.STE;

                    default:
                        return RefTagType.Unknown;
                }
            }

            public enum RefTagType { 
                Unknown, 
                A3D, 
                AF, 
                AR, 
                EL, 
                FP, 
                HD, 
                HP, 
                PL, 
                STC, 
                STE 
            }

            private static Dictionary<RefTagType, string> refTagTemplatePaths = new Dictionary<RefTagType, string>() {
                { RefTagType.Unknown,   @"\\REDACTED\sys\data\everyone\VDC Group\_BIM Support\Turing\Reftag Templates\Reftag_Default.dwg" },
                { RefTagType.A3D,       @"\\REDACTED\sys\data\everyone\VDC Group\_BIM Support\Turing\Reftag Templates\Reftag_A3D.dwg" },
                { RefTagType.AF,        @"\\REDACTED\sys\data\everyone\VDC Group\_BIM Support\Turing\Reftag Templates\Reftag_AF.dwg" },
                { RefTagType.AR,        @"\\REDACTED\sys\data\everyone\VDC Group\_BIM Support\Turing\Reftag Templates\Reftag_AR.dwg" },
                { RefTagType.EL,        @"\\REDACTED\sys\data\everyone\VDC Group\_BIM Support\Turing\Reftag Templates\Reftag_EL.dwg" },
                { RefTagType.FP,        @"\\REDACTED\sys\data\everyone\VDC Group\_BIM Support\Turing\Reftag Templates\Reftag_FP.dwg" },
                { RefTagType.HD,        @"\\REDACTED\sys\data\everyone\VDC Group\_BIM Support\Turing\Reftag Templates\Reftag_HD.dwg" },
                { RefTagType.HP,        @"\\REDACTED\sys\data\everyone\VDC Group\_BIM Support\Turing\Reftag Templates\Reftag_HP.dwg" },
                { RefTagType.PL,        @"\\REDACTED\sys\data\everyone\VDC Group\_BIM Support\Turing\Reftag Templates\Reftag_PL.dwg" },
                { RefTagType.STC,       @"\\REDACTED\sys\data\everyone\VDC Group\_BIM Support\Turing\Reftag Templates\Reftag_STC.dwg" },
                { RefTagType.STE,       @"\\REDACTED\sys\data\everyone\VDC Group\_BIM Support\Turing\Reftag Templates\Reftag_STE.dwg" }
            };

            public static FileInfo CreateRefTag(RefTagType type, string levelName, string scriptDestinationPath, string refTagFileDestinationPath) {
                // todo // set color if needed // reftag class?
                string tradeAbbrev;

                if (type == RefTagType.Unknown) {
                    Program.Log($"Reftag template not found for {Path.GetFileName(refTagFileDestinationPath)}, falling back to default reftag.", Program.MessageType.Warning);
                    tradeAbbrev = Regex.Match(refTagFileDestinationPath, @"Reftags\\(.+)-Lev").Groups[1].Value;
                }
                else {
                    tradeAbbrev = type.ToString();
                }

                var templatePath = refTagTemplatePaths[type];

                if (!File.Exists(templatePath)){
                    throw new FileNotFoundException($"Reftag template not found at: {templatePath}");
                }

                string fullScriptText;

                using (var reader = new StreamReader(@"\\REDACTED\sys\data\everyone\VDC Group\_BIM Support\Turing\Reftag Templates\edit-reftag-text.scr")) {
                    fullScriptText = reader.ReadToEnd();
                }

                fullScriptText =  fullScriptText.Replace("%TRADEABBREV%", tradeAbbrev)
                                                .Replace("%LEVEL%", levelName)
                                                .Replace("%FILEPATH%", refTagFileDestinationPath);

                File.WriteAllText(scriptDestinationPath, fullScriptText);

                RunACCC(templatePath, scriptDestinationPath, true); // saves to fileDestinationPath

                var fileInfo = new FileInfo(refTagFileDestinationPath);

                if (fileInfo.Exists) {
                    return fileInfo;
                }
                else {
                    throw new Exception($"RefTag could not be created: {refTagFileDestinationPath}.");
                }
            }
        }
    }
}