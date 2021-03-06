using CommandLine;

namespace ConformU
{
    public class CommandLineOptions
    {
        [Option('c', "commandline", Required = false, HelpText = "'--commandline' - Run Conform from the command line.")]
        public bool Run { get; set; }

        [Option('s', "settings", Required = false, HelpText = "'--settings fullyqualifiedfilename' - Fully qualified file name of the configuration file.")]
        public string SettingsFileLocation { get; set; }

        [Option('p', "logfilepath", Required = false, HelpText = "'--logfilepath fullyqualifiedlogfilepath' - Fully qualified path to the log file folder. Leave blank to use the default mechanic, which creates a new folder each day,")]
        public string LogFilePath { get; set; }

        [Option('n', "logfile", Required = false, HelpText = "'--logfile filename' - If filename has no directory/folder component it will be appended to the " +
            "log file path to create the fully qualified log file name. If filename is fully qualified, any logfilepath parameter will be ignored. Leave filename blank to use automatic file " +
            "naming, where the file name will be based on the file creation time.")]
        public string LogFileName { get; set; }

        [Option('r', "resultsfile", Required = false, HelpText = "'--resultsfile fullyqualifiedfilename' - Fully qualified file name of the results file.")]
        public string ResultsFileName { get; set; }

        [Option('d', "debugdiscovery", Required = false, HelpText = "'--debugdiscovery' - Write discovery debug information to the log.")]
        public bool DebugDiscovery { get; set; }

        [Option('t', "debugstartup", Required = false, HelpText = "'--debugstartup' - Write start-up debug information to the log.")]
        public bool DebugStartup { get; set; }

    }
}
