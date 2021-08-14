﻿using CommandLine;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ConformU
{
    public class CommandLineOptions
    {
        [Option('c', "commandline", Required = false, HelpText = "'--commandline' - Run Conform from the command line.")]
        public bool Run { get; set; }

        [Option('s', "settings", Required = false, HelpText = "'--settings fullyqualifiedfilename' - Fully qualified file name of the configuration file. Leave blank to use the default location.")]
        public string SettingsFileLocation { get; set; }

        [Option('p', "logfilepath", Required = false, HelpText = "'--logfilepath fullyqualifiedlogfilepath' - Fully qualified path to the log file folder. Leave blank to use the default mechanic, which creates a new folder each day,")]
        public string LogFilePath { get; set; }

        [Option('n', "logfilename", Required = false, HelpText = "'--logfilename filename' - If filename has no directory/folder component it will be appended to the " +
            "log file path to create the fully qualified log file name. If filename is fully qualified, any logfilepath parameter wil be ignored. Leave filename blank to use automatic file " +
            "naming, where the file name will be based on the file creation time.")]
        public string LogFileName { get; set; }
    }
}
