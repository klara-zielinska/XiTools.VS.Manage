//Inspired by http://stackoverflow.com/questions/350323/open-a-file-in-visual-studio-at-a-specific-line-number

using EnvDTE;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace XiTools.VS
{
    class CurrentDocPath
    {
        static void Main(string[] args)
        {
            try
            {
                if (args.Length > 1) 
                    Error("Invalid number of arguments", showHelp: true);

                string dteprocname = null;
                int pid = -1;
                if (args.Length == 1)
                {
                    dteprocname = TryParseVersionString(args[0]);
                    if (dteprocname == null)
                        if (!int.TryParse(args[0], out pid)) 
                            Error("Invalid parameter", showHelp: true);
                }

                var dte = GetDTEInstance(dteprocname, pid);
                if (dte == null) Error("VS instance not found", showHelp: true);
                Console.Out.Write(dte.ActiveDocument.FullName);
            }
            catch (Exception e)
            {
                Error(e.Message);
            }
        }

        static DTE GetDTEInstance(string dteprocname, int pid)
        {
            IRunningObjectTable rot;
            IEnumMoniker enumMoniker;

            if (dteprocname == null) dteprocname = "VisualStudio.DTE";

            if (GetRunningObjectTable(0, out rot) == 0)
            {
                rot.EnumRunning(out enumMoniker);

                var moniker = new IMoniker[1];
                while (enumMoniker.Next(1, moniker, new IntPtr()) == 0)
                {
                    CreateBindCtx(0, out var bindCtx);
                    moniker[0].GetDisplayName(bindCtx, null, out var displayName);

                    if (    displayName.StartsWith("!" + dteprocname) &&
                            (pid < 0 || displayName.EndsWith(":" + pid)) )
                    {
                        rot.GetObject(moniker[0], out var obj);
                        return obj as DTE;
                    }
                }
            }

            return null;
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
                "usage: {0} [<version>|<PID>]\n\n" +

                "version - either 'DTEx.y', where x.y is the version, or a supported version of Visual Studio.\n\n" +

                "Supported VS versions:\n\n",

                Path.GetFileName(Assembly.GetExecutingAssembly().Location));
            foreach (int version in supportedVersions)
                msg += String.Format("VS{0} ({1})\n", version, TryParseVersionString("VS" + version));

            return msg;
        }


        [DllImport("ole32.dll")]
        static extern void CreateBindCtx(int reserved, out IBindCtx ppbc);

        [DllImport("ole32.dll")]
        static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);
    }
}
