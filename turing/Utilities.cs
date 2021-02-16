using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace Turing {
    class Utilities
    {
        public static string MorningOrAfternoon() {
            var hour = DateTime.Now.Hour;
            if (hour < 11) {
                return "Morning";
            } else {
                return "Afternoon";
            }
        }

        public static string SwapQXDrive(string input) {
            string output;

            if (input.StartsWith(@"\\REDACTED\sys\data\everyone")) {
                output = input.Replace(@"\\REDACTED\sys\data\everyone", @"\\jcc-nas");
            }
            else if (input.StartsWith(@"\\redacted")) {
                output = input.Replace(@"\\redacted", @"\\REDACTED\sys\data\everyone");
            }
            else {
                throw new Exception("SwapQXDrive input did not start with \"\\\\REDACTED\\sys\\data\\everyone\" or \"jcc-nas\": " + input);
            }

            return output;
        }

        public static bool IsFileReady(string filename) {
            // If the file can be opened for exclusive access it means that the file
            // is no longer locked by another process.
            try {
                using (FileStream inputStream = File.Open(filename, FileMode.Open, FileAccess.Read, FileShare.None))
                    return inputStream.Length > 0;
            } catch (Exception) {
                return false;
            }
        }

        //public static void ZipFiles(string[] filePaths) {
        //    string timestamp = DateTime.Now.ToString("yyyy-MM-dd(HHmm)");
        //    string zipName = "backup_" + timestamp;
        //    string zipPath = Path.Combine(Path.GetDirectoryName(filePaths.First()), zipName) + ".zip";

        //    while (File.Exists(zipPath)) {
        //        zipName += "_1";
        //        zipPath = Path.Combine(Path.GetDirectoryName(filePaths.First()), zipName) + ".zip";
        //    }

        //    using (ZipArchive zipArchive = ZipFile.Open(zipPath, ZipArchiveMode.Create)) {
        //        foreach (string filePath in filePaths) {
        //            zipArchive.CreateEntryFromFile(filePath, Path.GetFileName(filePath));
        //        }
        //    }
        //}

        //public static void ZipFiles(string[] filePaths, string destinationDir) {
        //    string timestamp = DateTime.Now.ToString("yyyy-MM-dd(HHmm)");
        //    string zipName = "backup_" + timestamp;
        //    string zipPath = Path.Combine(destinationDir, zipName) + ".zip";

        //    while (File.Exists(zipPath)) {
        //        zipName += "_1";
        //        zipPath = Path.Combine(destinationDir, zipName) + ".zip";
        //    }

        //    using (ZipArchive zipArchive = ZipFile.Open(zipPath, ZipArchiveMode.Create)) {
        //        foreach (string filePath in filePaths) {
        //            zipArchive.CreateEntryFromFile(filePath, Path.GetFileName(filePath));
        //        }
        //    }
        //}

    }
}
