using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management;
using System.Text;

namespace Payload
{
    internal class Program
    {
        private static string printServer = @"\\DR3PRINT\";
        private static string useDirectory = @"C:\Printer_Migration\";
        private static string logFile = useDirectory + "printer.log";
        private static string errorLogFile = useDirectory + "error.log";
        private static List<string> printersToAdd, printersToRemove, installedPrinters, remove;
        private static ManagementScope oManagementScope = null;

        [System.Runtime.InteropServices.DllImport("winspool.drv")]
        public static extern int DeletePrinterConnection(string printerName);

        private static void Main(string[] args)
        {
            installedPrinters = new List<string>();
            printersToAdd = new List<string>();
            printersToRemove = new List<string>();
            remove = new List<string>();

            if (!Directory.Exists(useDirectory))
            {
                Directory.CreateDirectory(useDirectory);
            }
            validateCommand(args);
            DeleteRetiredPrinters();
            //AddNewPrinters();
        }

        private static void validateCommand(string[] args)
        {
            try
            {
                for (int i = 0; i < args.Length; i++)
                {
                    if (args[i].Equals("-a"))
                    {
                        while (!args[++i].Equals("-d") && i < args.Length)
                        {
                            printersToAdd.Add(args[i].Replace(',', ' ').Trim().ToUpper());
                        }
                    }
                    else if (args[i].Equals("-d"))
                    {
                        while (i < args.Length && !args[++i].Equals("-a"))
                        {
                            printersToRemove.Add(args[i].Replace(',', ' ').Trim().ToUpper());
                        }
                    }
                    i--;
                }
            }
            catch (System.IndexOutOfRangeException)
            {
                // Ok to be here, end of args
            }
        }

        private static void AddNewPrinters()
        {
            throw new NotImplementedException();
        }

        private static void DeleteRetiredPrinters()
        {
            GetInstalledPrinters();

            report("Installed Printers");
            report("=====================================");
            LogInstalledComputers();

            report("Printers needed to remove");
            report("=====================================");
            Compare();

            DeletePrinters();

            report("Installed Printers after Deletion");
            report("=====================================");
            GetInstalledPrinters();
            LogInstalledComputers();
        }

        private static void DeletePrinters()
        {
            if (remove.Count < 1)
                return;

            foreach (string printer in remove)
            {
                try
                {
                    if (DeletePrinterConnection(printServer + printer) == 0)
                    {
                        errorReport("Error removing printer " + printer);
                    }
                }
                catch (Exception ex)
                {
                    errorReport("Error removing printer " + printer + ". " + ex.ToString());
                }
            }
            System.Threading.Thread.Sleep(5000);
        }

        private static void Compare()
        {
            foreach (string printer in printersToRemove)
            {
                foreach (string installedPrinter in installedPrinters)
                {
                    if (printer.Equals(installedPrinters))
                    {
                        remove.Add(printer);
                        report(printer);
                    }
                }
            }
        }

        private static void GetInstalledPrinters()
        {
            RegistryKey key, subkey = null;
            try
            {
                key = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default);

                foreach (string value in key.GetSubKeyNames())
                {
                    // If printer connection exist save the user and list of printers
                    subkey = key.OpenSubKey(value + @"\Printers\Connections");
                    if (subkey == null)
                        continue;

                    foreach (string printer in subkey.GetSubKeyNames())
                    {
                        installedPrinters.Add(printer.Replace(',', ' ').Trim().ToUpper());
                    }
                }
                if (subkey != null)
                {
                    subkey.Close();
                }

                key.Close();
            }
            catch (Exception ex)
            {
                errorReport(ex.ToString());
            }
        }

        #region Logs

        private static void LogInstalledComputers()
        {
            foreach (string printer in installedPrinters)
            {
                report(printer);
                report("");
            }
        }

        private static void errorReport(string str)
        {
            System.IO.StreamWriter fstream = new System.IO.StreamWriter(errorLogFile, true);
            fstream.WriteLine(str);
            fstream.Close();
        }

        private static void report(string str)
        {
            System.IO.StreamWriter fstream = new System.IO.StreamWriter(logFile, true);
            fstream.WriteLine(str);
            fstream.Close();
        }

        #endregion Logs

        public static bool DeletePrinter(string sPrinterName)
        {
            oManagementScope = new ManagementScope(ManagementPath.DefaultPath);
            oManagementScope.Connect();

            SelectQuery oSelectQuery = new SelectQuery();
            oSelectQuery.QueryString = @"SELECT * FROM Win32_Printer WHERE Name = '" +
               sPrinterName.Replace("\\", "\\\\") + "'";

            ManagementObjectSearcher oObjectSearcher =
               new ManagementObjectSearcher(oManagementScope, oSelectQuery);
            ManagementObjectCollection oObjectCollection = oObjectSearcher.Get();

            if (oObjectCollection.Count != 0)
            {
                foreach (ManagementObject oItem in oObjectCollection)
                {
                    oItem.Delete();
                    return true;
                }
            }
            return false;
        }
    }
}