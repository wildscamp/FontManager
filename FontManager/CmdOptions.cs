using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CommandLine;
using CommandLine.Text;
using System.Reflection;

namespace Wilds.Apps.FontMgr
{
    class CmdOptions
    {
        [OptionArray('i', "install", Required = false, HelpText = "Install the given fonts.", DefaultValue = new string[] { })]
        public string[] InstallFiles { get; set; }


        [ValueList(typeof(List<string>))]
        public IList<string> ImplicitInstallFiles { get; set; }

        [OptionArray('u', "uninstall", Required = false, HelpText = "Uninstall the given fonts.", DefaultValue = new string[] { })]
        public string[] UninstallFiles { get; set; }

        [Option('q', DefaultValue = false, HelpText = "Suppress all output.")]
        public bool Quiet { get; set; }
        
        [HelpOption('?', "help", MutuallyExclusiveSet = "help")]
        public string GetUsage()
        {
            var help = new HelpText
            {
                Heading = new HeadingInfo("Font Manager", Assembly.GetEntryAssembly().GetName().Version.ToSt‌​ring()),
                AdditionalNewLineAfterOption = true,
                AddDashesToOption = true
            };
            help.AddPreOptionsLine(String.Format("{0}A simple utility for adding and removing system fonts in Windows.", Environment.NewLine));
            help.AddPreOptionsLine(String.Format("{0}Usage: {1} [<font file>* [-i <font file>*] [-u <font file>*]] [-q]", Environment.NewLine, Assembly.GetEntryAssembly().GetName().Name));
            help.AddOptions(this);
            return help;
        }
    }
}
