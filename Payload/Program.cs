using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Management;
using System.Net.NetworkInformation;
using System.Printing;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;

namespace Payload
{
    internal class Program
    {
        private static string printServer = @"\\DR3PRINT\";
        private static string printServer2 = @"\\DR3PRINT.ITSERVICES.NETWORK\";
        private static string useDirectory;
        private static string logFile;
        private static string errorLogFile;
        private static List<string> printersToAdd, printersToRemove, installedPrinters, remove;
        private static string driverPath = @"\\Iwmdocs\iwm\CIWMB-INFOTECH\Network\Printers\SHARP-MX-4141N\SOFTWARE-CDs\CD1\Drivers\Printer\English\PS\64bit\ss0hmenu.inf";
        private static string dbConnection = @"Data Source=W8-RKOEN;Initial Catalog=PrinterMigration;User ID=sa;Password=p@ssw0rd";
        private static bool successfulDelete, successfulAdd, error, defaultDeleted;
        private static string errorString, defaultPrinter;

        [System.Runtime.InteropServices.DllImport( "winspool.drv" )]
        public static extern int DeletePrinterConnection(string printerName);

        [System.Runtime.InteropServices.DllImport( "winspool.drv" )]
        public static extern int AddPrinterConnection(string printerName);

        [System.Runtime.InteropServices.DllImport( "winspool.drv" )]
        public static extern bool SetDefaultPrinter(string printerName);

        private static void Main(string[] args)
        {
            useDirectory = @"C:\Printer_Migration\";
            logFile = useDirectory + "printer.log";
            errorLogFile = useDirectory + "error.log";

            PrinterSettings settings = new PrinterSettings();
            foreach ( string name in PrinterSettings.InstalledPrinters )
            {
                settings.PrinterName = name;
                if ( settings.IsDefaultPrinter )
                {
                    defaultPrinter = name;
                }
            }
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
            generateTable();
        }

        private static void generateTable()
        {
            string user, removeList, deletion, addList, addition, status, errors;

            user = Environment.UserName;

            removeList = "";
            if ( remove.Count < 1 )
            {
                removeList = "Not Applicable";
            }
            else if ( remove.Count == 1 )
            {
                removeList += remove[0];
            }
            else
            {
                int counter = 1;
                foreach ( string printer in remove )
                {
                    if ( counter++ == remove.Count )
                        removeList += printer;
                    else
                        removeList += printer + ",";
                }
            }

            deletion = "";
            if ( remove.Count < 1 )
            {
                deletion = "Not Applicable";
            }
            else
            {
                if ( successfulDelete )
                {
                    deletion = "TRUE";
                }
                else
                {
                    deletion = "FALSE";
                }
            }

            addList = "";
            if ( printersToAdd.Count < 1 )
            {
                addList = "Not Applicable";
            }
            else if ( printersToAdd.Count == 1 )
            {
                addList += printersToAdd[0];
            }
            else
            {
                int counter = 1;
                foreach ( string printer in printersToAdd )
                {
                    if ( counter++ == printersToAdd.Count )
                        addList += printer;
                    else
                        addList += printer + ",";
                }
            }

            addition = "";
            if ( printersToAdd.Count < 1 )
            {
                addition = "Not Applicable";
            }
            else
            {
                if ( successfulAdd )
                {
                    addition = "TRUE";
                }
                else
                {
                    addition = "FALSE";
                }
            }

            status = "ONLINE";

            errors = "FALSE";
            if ( error )
            {
                errors = errorString;
            }

            using ( SqlConnection connection = new SqlConnection( dbConnection ) )
            {
                try
                {
                    connection.Open();
                }
                catch ( Exception ex )
                {
                    errorReport( "Couldnt open DB: " + ex.ToString() );
                }

                try
                {
                    using ( SqlCommand command = new SqlCommand( "INSERT INTO Main VALUES(@User, @RetiredPrinters, @SuccessfulDelete," +
                                                               "@AddedPrinters, @SuccesfulAdd, @Status, @Errors)", connection ) )
                    {
                        command.Parameters.Add( new SqlParameter( "User", user ) );
                        command.Parameters.Add( new SqlParameter( "RetiredPrinters", removeList ) );
                        command.Parameters.Add( new SqlParameter( "SuccessfulDelete", deletion ) );
                        command.Parameters.Add( new SqlParameter( "AddedPrinters", addList ) );
                        command.Parameters.Add( new SqlParameter( "SuccesfulAdd", addition ) );
                        command.Parameters.Add( new SqlParameter( "Status", status ) );
                        command.Parameters.Add( new SqlParameter( "Errors", errors ) );
                        command.ExecuteNonQuery();
                    }
                }
                catch ( Exception ex )
                {
                    errorReport( ex.ToString() );
                }
            }
        }

        #region ADD

        private static void AddNewPrinters()
        {
            if ( printersToAdd.Count < 1 )
                return;
            successfulAdd = true;

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
                        successfulAdd = false;
                        errorReport( "Error adding printer " + printer );
                    }
                    report( printer + " added" );

                    if ( defaultDeleted )
                    {
                        if ( !SetDefaultPrinter( printServer + printer ) )
                            report( "Setting " + printer + " to default failed!" );
                        else
                            report( printer + "set to default." );
                    }
                }
                catch ( Exception ex )
                {
                    successfulAdd = false;
                    errorReport( ex.ToString() + ". Error adding printer " + printer );
                    error = true;
                    errorString = ex.ToString() + ". Error adding printer " + printer;
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
                //installCertificate( Path.Combine( Environment.CurrentDirectory, Path.GetFileName( "SharpCert.cer" ) ) );
                installCertificate( @"C:\Users\Public\Documents\Printer_Migration\SharpCert.cer" );
                System.Diagnostics.Process proc = new System.Diagnostics.Process();
                System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
                startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                startInfo.FileName = "Rundll32.exe";
                startInfo.Arguments = @"printui.dll,PrintUIEntry /ia /f """ + driverPath + @"""  /m ""SHARP MX-4141N PS""";
                startInfo.Verb = "runas";
                startInfo.UseShellExecute = false;
                proc.StartInfo = startInfo;
                proc.Start();
                proc.WaitForExit();

                deleteCertificate();
            }
            catch ( Exception ex )
            {
                errorReport( ex.ToString() );
                error = true;
                errorString = ex.ToString() + ". Error installing print driver";
            }
        }

        private static void deleteCertificate()
        {
            X509Store store = new X509Store( StoreName.TrustedPublisher, StoreLocation.LocalMachine );
            store.Open( OpenFlags.ReadWrite );

            X509Certificate2Collection col = store.Certificates.Find( X509FindType.FindBySubjectName, "Sharp", false );

            foreach ( X509Certificate2 cert in col )
            {
                store.Remove( cert );
            }
            store.Close();
        }

        private static void installCertificate(string path)
        {
            X509Certificate2 certificate = new X509Certificate2( path );
            X509Store store = new X509Store( StoreName.TrustedPublisher, StoreLocation.LocalMachine );

            store.Open( OpenFlags.ReadWrite );
            store.Add( certificate );
            store.Close();
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
            if ( printersToRemove.Count < 1 )
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
            successfulDelete = true;
            defaultDeleted = false;
            foreach ( string printer in remove )
            {
                try
                {
                    if ( DeletePrinterConnection( printServer + printer ) == 0 )
                    {
                        if ( DeletePrinterConnection( printServer2 + printer ) == 0 )
                        {
                            successfulDelete = false;
                            errorReport( "Error removing printer " + printer );
                        }
                    }

                    if ( successfulDelete )
                    {
                        if ( defaultPrinter.Equals( printServer + printer, StringComparison.InvariantCultureIgnoreCase ) ||
                           defaultPrinter.Equals( printServer2 + printer, StringComparison.InvariantCultureIgnoreCase ) )
                        {
                            report( "Default Printer deleted!" );
                            defaultDeleted = true;
                        }
                    }
                }
                catch ( Exception ex )
                {
                    errorReport( "Error removing printer " + printer + ". " + ex.ToString() );
                    successfulDelete = false;
                    error = true;
                    errorString = "Error removing printer " + printer + ". " + ex.ToString();
                }
                System.Threading.Thread.Sleep( 1000 );
            }
        }

        private static void Compare()
        {
            Regex r;

            foreach ( string printer in printersToRemove )
            {
                foreach ( string installedPrinter in installedPrinters )
                {
                    r = new Regex( "(" + printer + ")", RegexOptions.IgnoreCase );
                    if ( r.IsMatch( installedPrinter ) )
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
            error = true;
            errorString = str;

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
    }
}