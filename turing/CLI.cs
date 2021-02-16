using CommandLine;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Turing {
    public class CLI {
        [Verb("update")]
        public class UpdateOptions {
            [Option('p', "project", Required = true, HelpText = "Project number, e.g. '555555'")]
            public string Project { get; set; }

            [Option("nodl", Default = false, HelpText = "Do not start download from 360Sync")]
            public bool Nodl { get; set; }

            [Option("nwd", Default = false, HelpText = "Create NWDs of all newly updated CAD files")]
            public bool Nwd { get; set; }

            [Option('u', "show-unknown", Default = false, HelpText = "List unknown files in log")]
            public bool ShowUnknown { get; set; }

            [Option('i', "show-ignored", Default = false, HelpText = "List ignored files in log")]
            public bool ShowIgnored { get; set; }
        }
    }
}
