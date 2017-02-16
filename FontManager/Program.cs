using Microsoft.Win32;
using NLog;
using System;
using System.Collections.Generic;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace TheWilds
{

    /// <summary>
    /// Font Manager is a simple program that simplifies the ability to add and remove
    /// fonts for a Windows system from a command line. It has the benefit of being very
    /// fast in comparison to using the Windows Shell COM object.
    /// 
    /// The main ideas from this code were borrowed from KYRIACOSS at http://pastebin.com/C99TmXBn.
    /// This code pasting was referenced on StackOverflow at http://stackoverflow.com/a/29334492/277798.
    /// </summary>
    class FontManager
    {
        /// <summary>
        /// The path in HKLM where the fonts are registered.
        /// </summary>
        private const string FONT_REG_PATH = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts";

        /// <summary>
        /// Used to tell SendMessage() to broadcast that a font change has occurred.
        /// </summary>
        private const int WM_FONTCHANGE = 0x1D;

        /// <summary>
        /// Used to tell SendMessage() to broadcast to all top-level applications that a change has occurred.
        /// </summary>
        private const int HWND_BROADCAST = 0xffff;

        /// <summary>
        /// Notifies the the Windows graphics system that a font has been added.
        /// </summary>
        /// <param name="lpszFilename">A string that names a font resource file. Use full path name. To add a font whose information comes from several resource files, have lpszFileName point to a string with the file names separated by a "|" --for example, abcxxxxx.pfm | abcxxxxx.pfb.</param>
        /// <returns>If the function succeeds, the return value specifies the number of fonts added. If the function fails, the return value is zero. No extended error information is available.</returns>
        [DllImport("gdi32.dll")]
        private static extern int AddFontResource(string lpszFilename);

        /// <summary>
        /// Notifies the the Windows graphics system that a font has been removed.
        /// </summary>
        /// <param name="lpszFilename">A string that names a font resource file. Use full path name.</param>
        /// <returns></returns>
        [DllImport("gdi32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool RemoveFontResource(string lpszFilename);

        /// <summary>
        /// Notify other applications on the system when a change has happened. In the case
        /// of this program it notifies them when a font has been added or removed.
        /// </summary>
        /// <param name="hWnd">A handle to the window whose window procedure will receive the message. If this parameter is HWND_BROADCAST ((HWND)0xffff), the message is sent to all top-level windows in the system, including disabled or invisible unowned windows, overlapped windows, and pop-up windows; but the message is not sent to child windows.</param>
        /// <param name="Msg">The message to be sent. For lists of the system-provided messages, see ( System-Defined Messages | https://msdn.microsoft.com/en-us/library/windows/desktop/ms644927(v=vs.85).aspx#system_defined ) .</param>
        /// <param name="wParam">Additional message-specific information.</param>
        /// <param name="lParam">Additional message-specific information.</param>
        /// <returns>The return value specifies the result of the message processing; it depends on the message sent.</returns>
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);

        /// <summary>
        /// The logger is used for handling Console output as well as debug information.
        /// </summary>
        private static Logger logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// The main entrypoint for the application.
        /// </summary>
        /// <param name="args">All command-line arguments.</param>
        static void Main(string[] args)
        {
            ConfigureLogger();
            var options = new CmdOptions();

            if (CommandLine.Parser.Default.ParseArguments(args, options))
            {
                if (options.Quiet)
                {
                    // Remove all Console logging targets.
                    NLog.Config.LoggingRule[] rules = LogManager.Configuration.LoggingRules.Where(r => r.Targets.Where(t => t is NLog.Targets.ConsoleTarget).Count() > 0).ToArray();
                    foreach (var rule in rules)
                    {
                        LogManager.Configuration.LoggingRules.Remove(rule);
                    }

                    LogManager.ReconfigExistingLoggers();
                }

                IEnumerable<string> installFiles = options.InstallFiles.Concat(options.ImplicitInstallFiles);
                foreach (string file in installFiles)
                {
                    string FullPath = "";

                    // The path string may not represent a wellformed path.
                    try
                    {
                        FullPath = Path.GetFullPath(file);
                    }
                    catch (Exception)
                    {
                        logger.Warn("'{0}' is an malformed path. Unable to install.", file);

                        continue;
                    }

                    logger.Debug("Full path of file to install: '{0}'", FullPath);

                    if (!File.Exists(FullPath))
                    {
                        logger.Warn("'{0}' does not exist. Unable to install.", FullPath);
                    } else
                    {
                        if (InstallFont(FullPath))
                        {
                            logger.Info("Installed '{0}'", Path.GetFileName(FullPath));
                        }
                    }
                }

                foreach (string file in options.UninstallFiles)
                {
                    if (UninstallFont(file))
                    {
                        logger.Info("Uninstalled '{0}'", Path.GetFileName(file));
                    }
                }
            }
        }

        /// <summary>
        /// Define defaults for logger if there is no config file present.
        /// </summary>
        private static void ConfigureLogger()
        {
            if (!File.Exists(Path.Combine(Directory.GetCurrentDirectory(), "NLog.config")))
            {
                NLog.Targets.ConsoleTarget console = new NLog.Targets.ConsoleTarget("WriteLine");
                console.Layout = new NLog.Layouts.SimpleLayout("${message}");
                LogManager.Configuration = new NLog.Config.LoggingConfiguration();
                LogManager.Configuration.AddRule(LogLevel.Info, LogLevel.Fatal, console);
                LogManager.ReconfigExistingLoggers();
            }
        }


        /// <summary>
        /// Installs the specified font on the system.
        /// </summary>
        /// <param name="FontPath">The path to the font file to install. This could be an absolute
        /// path or a relative path from the application.</param>
        /// <returns>True if the font was installed. False if it was not installed.</returns>
        private static bool InstallFont(string FontPath)
        {
            bool installed = false;

            string FontName = Path.GetFileName(FontPath);

            // Creates the full path where your font will be installed
            var fontDestination = Path.Combine(System.Environment.GetFolderPath
                                          (System.Environment.SpecialFolder.Fonts), FontName);

            logger.Debug("Attempting to install '{0}' into '{1}'", FontName, fontDestination);

            if (!File.Exists(fontDestination))
            {
                logger.Debug("Copying '{0}' to '{1}'", FontPath, fontDestination);

                // Copies font to destination
                try
                {
                    System.IO.File.Copy(FontPath, fontDestination);
                }
                catch (Exception ex)
                {
                    logger.Error(ex.Message);

                    return false;
                }

                string actualFontName = GetFontName(fontDestination);

                logger.Debug(@"Setting 'HKLM\{0} [{1}]' = '{2}'.", FONT_REG_PATH, actualFontName, FontName);

                //Add registry entry
                Registry.SetValue(String.Format(@"HKEY_LOCAL_MACHINE\{0}", FONT_REG_PATH), actualFontName, FontName, RegistryValueKind.String);

                logger.Debug("Calling gdi32.dll::AddFontResource('{0}')", fontDestination);

                //Add font
                AddFontResource(fontDestination);

                logger.Debug("Using user32.dll::SendMessage() to notify other apps that a font change has happend.");

                SendMessage(new IntPtr(HWND_BROADCAST), WM_FONTCHANGE, IntPtr.Zero, IntPtr.Zero);

                installed = true;
            } else
            {
                logger.Debug("Did not install the font because it already exists in '{0}'", fontDestination);
            }

            return installed;
        }

        private static string GetFontName(string fontDestination)
        {
            logger.Debug("Attempting to retrieve font name from '{0}'", fontDestination);

            // Retrieves font name
            PrivateFontCollection fontCol = new PrivateFontCollection();
            fontCol.AddFontFile(fontDestination);

            var actualFontName = Path.GetFileName(fontDestination);

            if (fontCol.Families.Count() > 0)
            {
                actualFontName = fontCol.Families[0].Name;

                string extension = Path.GetExtension(fontDestination);
                
                switch (extension.ToLower())
                {
                    case ".ttf":
                    case ".ttc":
                        logger.Debug("Font type is TrueType. Appending '(TrueType)' to the end of the name.");

                        actualFontName = String.Format("{0} (TrueType)", actualFontName);
                        break;
                    default:
                        break;
                }
            } else
            {
                logger.Debug("Unable to retrive the font name from '{0}'. Using the file name instead.", fontDestination);
            }

            logger.Debug("Retrieved font name is '{0}'", actualFontName);

            return actualFontName;
        }

        /// <summary>
        /// Uninstall a font with the given file name.
        /// </summary>
        /// <param name="FileName">The name of the font file to uninstall. This could be just a file name
        /// or a full path to a font file. If it is a full path, it will extract the filename from the
        /// path and attempt to uninstall a font with that file name from the system's font folder.</param>
        /// <returns>True if the font was uninstalled. False if it was not.</returns>
        private static bool UninstallFont(string FileName)
        {
            bool uninstalled = false;

            string FontFileName = Path.GetFileName(FileName);

            // Creates the full path where your font will be uninstalled
            var systemFontPath = Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Fonts), FontFileName);

            logger.Debug("Attempting to uninstall '{0}'", systemFontPath);

            if (File.Exists(systemFontPath))
            {

                logger.Debug("Calling gdi32.dll::RemoveFontResource('{0}')", systemFontPath);

                bool result = RemoveFontResource(systemFontPath);


                logger.Debug("'{0}' does exist. Attempting to delete.");

                try
                {
                    System.IO.File.Delete(systemFontPath);
                }
                catch (UnauthorizedAccessException ex)
                {
                    logger.Error(ex.Message);

                    return false;
                }

                logger.Debug(@"Searching for any keys in 'HKLM\{0}' that reference '{1}' so they can be deleted.", FONT_REG_PATH, FontFileName);

                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(FONT_REG_PATH, true))
                {
                    if (key != null)
                    {
                        string[] names = key.GetValueNames();

                        // Delete all registry keys that point to the font file.
                        foreach (string name in names)
                        {
                            if (key.GetValueKind(name) == RegistryValueKind.String)
                            {
                                string namevalue = (string)key.GetValue(name);

                                if (namevalue.ToLower().Trim() == FontFileName.ToLower().Trim())
                                {
                                    logger.Debug(@"Deleting 'HKLM\{0} [{1}]' because it was referencing '{2}'", FONT_REG_PATH, name, FontFileName);

                                    key.DeleteValue(name);
                                }
                            }
                            
                        }
                    }
                }

                logger.Debug("Using user32.dll::SendMessage() to notify other apps that a font change has happened.");

                SendMessage(new IntPtr(HWND_BROADCAST), WM_FONTCHANGE, IntPtr.Zero, IntPtr.Zero);

                uninstalled = true;
            } else
            {
                logger.Debug("Did not uninstall '{0}' because it did not exist on the system.", systemFontPath);
            }
            
            return uninstalled;
        }
    }
}
