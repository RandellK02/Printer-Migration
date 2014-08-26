using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management;
using System.Net.NetworkInformation;
using System.Printing;
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
        private static string driverPath = @"\\Iwmdocs\iwm\CIWMB-INFOTECH\Network\Printers\SHARP-MX-4141N\SOFTWARE-CDs\CD1\Drivers\Printer\English\PS\64bit\ss0hmenu.inf";
        //private static ManagementScope oManagementScope = null;

        [System.Runtime.InteropServices.DllImport( "winspool.drv" )]
        public static extern int DeletePrinterConnection(string printerName);

        [System.Runtime.InteropServices.DllImport( "winspool.drv" )]
        public static extern int AddPrinterConnection(string printerName);

        [System.Runtime.InteropServices.DllImport( "Setupapi.dll", EntryPoint = "InstallHinfSection", CallingConvention = System.Runtime.InteropServices.CallingConvention.StdCall )]
        public static extern void InstallHinfSection(
            IntPtr hwnd,
            IntPtr ModuleHandle,
            string CmdLineBuffer,
            int nCmdShow);

        private static void Main(string[] args)
        {
            installedPrinters = new List<string>();
            printersToAdd = new List<string>();
            printersToRemove = new List<string>();
            remove = new List<string>();

            if ( !Directory.Exists( useDirectory ) )
            {
                Directory.CreateDirectory( useDirectory );
            }
            validateCommand( args );
            DeleteRetiredPrinters();
            AddNewPrinters();

            report( "END OF REPORT" );
            errorReport( "END OF REPORT" );
        }

        #region ADD

        private static void AddNewPrinters()
        {
            if ( printersToAdd.Count < 1 )
                return;

            report( "Adding New Printer" );
            report( "=====================================" );

            foreach ( string printer in printersToAdd )
            {
                try
                {
                    if ( !validatePrinterName( printer ) )
                    {
                        report( printer + " Not Found!" );
                        continue;
                    }

                    // Install Print Driver
                    installPrinterDriver( driverPath );

                    if ( AddPrinterConnection( printServer + printer ) == 0 )
                    {
                        errorReport( "Error adding printer " + printer );
                    }
                    report( printer + " added" );
                }
                catch ( Exception ex )
                {
                    errorReport( ex.ToString() + ". Error adding printer " + printer );
                }
            }

            report( "" );
            report( "Installed Printers after Additions" );
            report( "=====================================" );
            installedPrinters.Clear();
            GetInstalledPrinters();
            LogInstalledComputers();
        }

        private static void installPrinterDriver(string driverPath)
        {
            try
            {
                InstallHinfSection( IntPtr.Zero, IntPtr.Zero, driverPath, 0 );
                /*System.Diagnostics.Process proc = new System.Diagnostics.Process();
                System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
                startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                startInfo.FileName = "Rundll32.exe";
                startInfo.Arguments = @"advpack.dll,LauchINFSectionEx """ + driverPath + @""" /m ""Sharp""";
                startInfo.Verb = "runas";
                startInfo.UseShellExecute = false;
                proc.StartInfo = startInfo;
                proc.Start();
                proc.WaitForExit();*/
            }
            catch ( Exception ex )
            {
                errorReport( ex.ToString() );
            }
        }

        private static bool validatePrinterName(string printerName)
        {
            bool found = false;
            if ( !ping( printerName ) )
            {
                using ( PrintServer printSrvr = new PrintServer( string.Format( @"\\{0}", "DR3PRINT" ) ) )
                {
                    foreach ( var printer in printSrvr.GetPrintQueues() )
                    {
                        if ( printer.Name.ToUpper().Equals( printerName.ToUpper() ) )
                        {
                            found = true;
                            break;
                        }
                    }
                }
            }
            else
            {
                found = true;
            }

            return found;
        }

        #endregion ADD

        #region DELETE

        private static void DeleteRetiredPrinters()
        {
            if ( remove.Count < 1 )
                return;
            GetInstalledPrinters();

            report( "Installed Printers" );
            report( "=====================================" );
            LogInstalledComputers();

            report( "Printers needed to remove" );
            report( "=====================================" );
            Compare();
            report( "" );

            DeletePrinters();

            report( "Installed Printers after Deletion" );
            report( "=====================================" );
            installedPrinters.Clear();
            GetInstalledPrinters();
            LogInstalledComputers();
        }

        private static void DeletePrinters()
        {
            foreach ( string printer in remove )
            {
                try
                {
                    if ( DeletePrinterConnection( printServer + printer ) == 0 )
                    {
                        errorReport( "Error removing printer " + printer );
                    }
                }
                catch ( Exception ex )
                {
                    errorReport( "Error removing printer " + printer + ". " + ex.ToString() );
                }
                System.Threading.Thread.Sleep( 1000 );
            }
        }

        private static void Compare()
        {
            foreach ( string printer in printersToRemove )
            {
                foreach ( string installedPrinter in installedPrinters )
                {
                    if ( printer.Equals( installedPrinter ) )
                    {
                        remove.Add( printer );
                        report( printer );
                    }
                }
            }
        }

        #endregion DELETE

        private static void GetInstalledPrinters()
        {
            RegistryKey key;
            string temp;
            try
            {
                key = Registry.CurrentUser.OpenSubKey( @"Printers\Connections" );
                foreach ( string printer in key.GetSubKeyNames() )
                {
                    temp = printer.Replace( ',', ' ' ).Trim().ToUpper();
                    temp = temp.Replace( "DR3PRINT", "" ).Trim();

                    installedPrinters.Add( temp );
                }

                key.Close();
            }
            catch ( Exception ex )
            {
                errorReport( ex.ToString() );
            }
        }

        private static bool ping(string server)
        {
            Ping p = new Ping();
            try
            {
                PingReply reply = p.Send( server );
                if ( reply.Status == IPStatus.Success )
                    return true;
            }
            catch { }
            return false;
        }

        private static void validateCommand(string[] args)
        {
            try
            {
                for ( int i = 0; i < args.Length; i++ )
                {
                    if ( args[i].Equals( "-a" ) )
                    {
                        while ( !args[++i].Equals( "-d" ) && i < args.Length )
                        {
                            printersToAdd.Add( args[i].Replace( ',', ' ' ).Trim().ToUpper() );
                        }
                    }
                    else if ( args[i].Equals( "-d" ) )
                    {
                        while ( i < args.Length && !args[++i].Equals( "-a" ) )
                        {
                            printersToRemove.Add( args[i].Replace( ',', ' ' ).Trim().ToUpper() );
                        }
                    }
                    i--;
                }
            }
            catch ( System.IndexOutOfRangeException )
            {
                // Ok to be here, end of args
            }
        }

        #region Logs

        private static void LogInstalledComputers()
        {
            foreach ( string printer in installedPrinters )
            {
                report( printer );
            }
            report( "" );
        }

        private static void errorReport(string str)
        {
            System.IO.StreamWriter fstream = new System.IO.StreamWriter( errorLogFile, true );
            fstream.WriteLine( str );
            fstream.Close();
        }

        private static void report(string str)
        {
            System.IO.StreamWriter fstream = new System.IO.StreamWriter( logFile, true );
            fstream.WriteLine( str );
            fstream.Close();
        }

        #endregion Logs

        /*
        public static bool DeletePrinter(string sPrinterName)
        {
            oManagementScope = new ManagementScope( ManagementPath.DefaultPath );
            oManagementScope.Connect();

            SelectQuery oSelectQuery = new SelectQuery();
            oSelectQuery.QueryString = @"SELECT * FROM Win32_Printer WHERE Name = '" +
               sPrinterName.Replace( "\\", "\\\\" ) + "'";

            ManagementObjectSearcher oObjectSearcher =
               new ManagementObjectSearcher( oManagementScope, oSelectQuery );
            ManagementObjectCollection oObjectCollection = oObjectSearcher.Get();

            if ( oObjectCollection.Count != 0 )
            {
                foreach ( ManagementObject oItem in oObjectCollection )
                {
                    oItem.Delete();
                    return true;
                }
            }
            return false;
        }*/
    }
}