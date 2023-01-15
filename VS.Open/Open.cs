//Inspired by http://stackoverflow.com/questions/350323/open-a-file-in-visual-studio-at-a-specific-line-number

using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace XiTools.VS
{
	class Open
	{
		static void Main(string[] args)
		{
			try
			{
				if (args.Length != 2 && args.Length != 3)
                    Error("Invalid number of arguments", showHelp: true);

				string vsString = TryParseVersionString(args[0]);
				if (vsString == null) 
                    Error("Usupported version", showHelp: true);

                string filename = args[1];
                int line = -1;
				if ( args.Length == 3 && (!int.TryParse(args[2], out line) || line < 0) )
                    Error("Invalid line number");

                var dte2 = Marshal.GetActiveObject(vsString) as EnvDTE80.DTE2;
				dte2.MainWindow.Activate();
				EnvDTE.Window w = dte2.ItemOperations.OpenFile(filename, EnvDTE.Constants.vsViewKindTextView);
				if (line != -1)
				{
					for (int i = 0; i < 200; ++i)
						try
						{
							System.Threading.Thread.Sleep(50);
							((EnvDTE.TextSelection)dte2.ActiveDocument.Selection).GotoLine(line, true);
							break;
						}
						catch (COMException) { }
				}
			}
			catch(Exception e)
			{
                Error(e.Message, false);
            }
        }

        static string TryParseVersionString(string visualOrDTEVersion)
        {
            if (visualOrDTEVersion.StartsWith("VS"))
                switch (visualOrDTEVersion)
                {
                    case "VS2022":
                        return "VisualStudio.DTE.17.0";
                    case "VS2019":
                        return "VisualStudio.DTE.16.0"; // WARRNING: It is a guess
                    case "VS2013":
                        return "VisualStudio.DTE.12.0";
                    case "VS2012":
                        return "VisualStudio.DTE.11.0";
                    case "VS2010":
                        return "VisualStudio.DTE.10.0";
                    case "VS2008":
                        return "VisualStudio.DTE.9.0";
                    case "VS2005":
                        return "VisualStudio.DTE.8.0";
                    case "VS2003":
                        return "VisualStudio.DTE.7.1";
                    case "VS2002":
                        return "VisualStudio.DTE.7";
                }

            else if (visualOrDTEVersion.StartsWith("DTE"))
                return "VisualStudio.DTE." + visualOrDTEVersion.Substring(3);

            return null;
        }

        static void Error(string msg, bool showHelp = false)
        {
            Console.Error.WriteLine("Error: " + msg + (showHelp ? "\n\n---\n" + GetHelpMessage() : ""));
            Environment.Exit(-1);
        }

        static string GetHelpMessage()
        {
            var supportedVersions = new List<int>() { 2002, 2003, 2005, 2008, 2010, 2012, 2013, 2019, 2022 };

            var msg = String.Format(
                "usage: {0} <version> <file path> [<line number>]\n\n" +

                "version - either 'DTEx.y', where x.y is the version, or a supported version of Visual Studio.\n\n" +

                "Supported VS versions:\n\n",

                Path.GetFileName(Assembly.GetExecutingAssembly().Location));
            foreach (int version in supportedVersions)
                msg += String.Format("VS{0} ({1})\n", version, TryParseVersionString("VS" + version));

            return msg;
        }
    }
}
