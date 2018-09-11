using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ExcelToCsvConverter
{
    class Program
    {
        #region// Unmanaged

        // Declare the SetConsoleCtrlHandler function
        // as external and receiving a delegate.
        [DllImport("Kernel32")]
        public static extern bool SetConsoleCtrlHandler(HandlerRoutine Handler, bool Add);

        // A delegate type to be used as the handler routine
        // for SetConsoleCtrlHandler.
        public delegate bool HandlerRoutine(CtrlTypes CtrlType);

        public enum CtrlTypes
        {
            CTRL_C_EVENT = 0,
            CTRL_BREAK_EVENT = 1,
            CTRL_CLOSE_EVENT = 2,
            CTRL_LOGOFF_EVENT = 5,
            CTRL_SHUTDOWN_EVENT = 6
        }

        #endregion

        private static readonly string ExcelResource = "Excel";

        private static Settings fSettings = new Settings();
        internal static Settings Settings
        {
            get { return fSettings; }
            set { fSettings = value; }
        }


        private static bool LoadSettings()
        {
            try
            {
                string settFile = "config" + ".xml";
                if (!File.Exists(settFile))
                {
                    Settings.SaveToXml(settFile, Settings);
                    Console.WriteLine("Settings file not found!\nCreate a default settings file.");
                    return false;
                }
                else
                {
                    Settings = Settings.LoadFromXml(settFile);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Can't load config file : " + e.Message);
                return false;
            }

            return true;
        }

        private static void ShowHelp()
        {
            var exName = Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().Location);
            Console.WriteLine("Convert from Excel to CSV: \t" + exName + " -csv [optional]");
            Console.WriteLine("Convert from CSV to Excel: \t" + exName + " -exc");
            Console.WriteLine("Help: \t\t\t\t" + exName + " /?");
            Console.WriteLine();
        }


        static ExcelCsvConverter converter;
        static void Main(string[] args)
        {
            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
            SetConsoleCtrlHandler(new HandlerRoutine(ConsoleCtrlCheck), true);

            LoadSettings();

            if (args.Length > 1)
            {
                Console.WriteLine("Invalid parameters count");
                return;
            }
            else
            {
                converter = new ExcelCsvConverter(Settings);

                if (args.Contains("?") || args.Contains("help")) ShowHelp();
                else if (args.Contains("-exc")) converter.CSVToExcel();
                //else if (args.Contains("-csv"))
                else converter.ExcelToCsv();

                converter.Dispose();
            }
        }

        private static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            var assemblyName = new AssemblyName(args.Name).Name;
            if (assemblyName == ExcelResource)
            {
                var allRes = typeof(Program).Assembly.GetManifestResourceNames();
                var resource = allRes.FirstOrDefault(line => line.Contains(ExcelResource));

                if (!String.IsNullOrEmpty(resource))
                {
                    using (Stream stream = typeof(Program).Assembly.GetManifestResourceStream(resource))
                    {
                        byte[] data = new BinaryReader(stream).ReadBytes((int)stream.Length);
                        return Assembly.Load(data);
                    }
                }
            }

            return null;
        }

        private static bool ConsoleCtrlCheck(CtrlTypes ctrlType)
        {
            switch (ctrlType)
            {
                case CtrlTypes.CTRL_C_EVENT:
                case CtrlTypes.CTRL_BREAK_EVENT:
                case CtrlTypes.CTRL_CLOSE_EVENT:
                case CtrlTypes.CTRL_LOGOFF_EVENT:
                case CtrlTypes.CTRL_SHUTDOWN_EVENT:
                    if (converter != null) converter.Dispose();
                    break;
            }

            return true;
        }
    }
}