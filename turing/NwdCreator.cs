using Autodesk.Navisworks.Api.Automation;
using System;
using System.Diagnostics;
using System.IO;

namespace Turing
{
    public class NwdCreator
    {
        static string navisOptionsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "navisoptions.xml");

        /// <summary>
        /// Takes input file, converts to NWD, and saves to output file. Overwrites if necessary.
        /// </summary>
        /// <param name="input">File path of any Navisworks compatible file type.</param>
        /// <param name="output">File path of output. File extension should be nwd.</param>
        public static void Create(string input, string output) {
            var processInfo = new ProcessStartInfo() {
                FileName = Program.turingSettings.RoamerPath,
                Arguments = $"-nogui -options \"{navisOptionsPath}\" -exit -nwd \"{output}\" \"{input}\""
            };
            var process = new Process() {
                StartInfo = processInfo
            };
            process.Start();
            process.WaitForExit(5 * 60 * 1000);
        }
    }
}
