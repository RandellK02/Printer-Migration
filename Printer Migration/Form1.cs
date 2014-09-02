using HTMLReportEngine;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Printing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Printer_Migration
{
    public partial class Form1 : Form
    {
        private static string useDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) + @"\Printer_Migration\";
        private static string logFile = useDirectory + "printer.log";
        private static string errorLogFile = useDirectory + "error.log";
        private static DataSet set = null;
        private static PrintServer server;
        private static PrintQueueCollection printersCollection;

        public Form1()
        {
            InitializeComponent();
            set = new DataSet("Printer DataSet");
            if (!Directory.Exists(useDirectory))
            {
                Directory.CreateDirectory(useDirectory);
            }
            server = new PrintServer(@"\\DR3PRINT");
            printersCollection = server.GetPrintQueues();

            foreach (PrintQueue printer in printersCollection)
            {
                Globals.printersOnServer.Add(printer.Name);
            }

            AutoCompleteStringCollection source = new AutoCompleteStringCollection();
            source.AddRange(Globals.printersOnServer.ToArray());
            tbAddPrinters.AutoCompleteCustomSource = source;
            tbDeletePrinters.AutoCompleteCustomSource = source;
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    Globals.userXLSPath = openFileDialog1.FileName;
                    if (!bwUserImport.IsBusy)
                    {
                        bwUserImport.RunWorkerAsync();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }

        #region Import

        private void bwUserImport_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            int progress = 0;
            double realProgress = 0;
            Excel.Workbook xlWorkBook = null;

            try
            {
                // Get users from excel worksheet
                Excel.Application xlApp = new Excel.Application();
                xlApp.Visible = false;
                xlWorkBook = xlApp.Workbooks.Open(Globals.userXLSPath);
                Excel.Worksheet xlWorkSheet = xlWorkBook.Sheets[1];
                int lastRow = xlWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                object[,] MyValues = xlWorkSheet.get_Range("A1", "A" + lastRow.ToString()).Value;

                string user;
                foreach (object value in MyValues)
                {
                    user = value.ToString();
                    Globals.users.Add(user);
                    realProgress += (100f / lastRow);
                    progress = (int)Math.Round(realProgress);
                    worker.ReportProgress(progress);
                }

                xlWorkBook.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                xlWorkBook.Close();
            }

            #region OLD

            /* Object thisLock = new Object();
         int myStreamLength = (int)myStream.Length;
         int cur;
         int progress = 0;
         double realProgress = 0;
         string btyesToString = "", modifiedBytesToString = "";

         lock ( thisLock )
         {
             using ( myStream )
             {
                 while ( (cur = myStream.ReadByte()) != -1 )
                 {
                     if ( cur != 44 )
                     {
                         Globals.temp.Add( (byte)cur );
                     }
                     else
                     {
                         btyesToString = System.Text.Encoding.Default.GetString( Globals.temp.ToArray() );
                         modifiedBytesToString = btyesToString.Replace( System.Environment.NewLine, "" );
                         Globals.users.Add( modifiedBytesToString );
                         Globals.temp.Clear();
                     }

                     realProgress += (100f / myStreamLength);
                     progress = (int)Math.Round( realProgress );
                     worker.ReportProgress( progress );
                 }

                 // Get last user if any
                 btyesToString = System.Text.Encoding.Default.GetString( Globals.temp.ToArray() );
                 modifiedBytesToString = btyesToString.Replace( System.Environment.NewLine, "" );
                 Globals.users.Add( modifiedBytesToString );

                 Globals.temp.Clear();
                 Globals.users.RemoveAll( string.IsNullOrWhiteSpace );
                 worker.ReportProgress( 100 );
             }
         }*/

            #endregion OLD
        }

        private void bwUserImport_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            pbUserImport.Value = e.ProgressPercentage;
        }

        #endregion Import

        private void btnRun_Click(object sender, EventArgs e)
        {
            if (bwRun.IsBusy != true)
                bwRun.RunWorkerAsync();
        }

        #region bwRun

        private void bwRun_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            string what = e.UserState.ToString();
            switch (e.UserState.ToString())
            {
                case "GENERATECOMPUTERNAMES":
                    pbGenComputers.Value = e.ProgressPercentage;
                    break;

                case "PINGCOMPUTERS":
                    pbPingComputers.Value = e.ProgressPercentage;
                    break;

                case "TRANSFERPAYLOAD":
                    pbSendPayload.Value = e.ProgressPercentage;
                    break;

                case "EXECUTEPAYLOAD":
                    pbExecPayload.Value = e.ProgressPercentage;
                    break;

                case "GENERATEREPORT":
                    pbReport.Value = e.ProgressPercentage;
                    break;

                case "DELETEPAYLOAD":
                    pbCleanUp.Value = e.ProgressPercentage;
                    break;
            }
        }

        private void bwRun_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            generateComputerNames(worker);
            pingComputers(worker);
            System.Threading.Thread.Sleep(1000);
            transferPayload(worker);
            System.Threading.Thread.Sleep(1000);
            executePayload(worker);
            System.Threading.Thread.Sleep(5000);
            generateReport(worker);
            System.Threading.Thread.Sleep(1000);
            removePayload(worker);

            report("Offline Computers");
            report("=====================================");
            foreach (string computer in Globals.offlineComputers)
            {
                report(computer);
            }
        }

        private void removePayload(BackgroundWorker worker)
        {
            int progress = 0;
            double realProgress = 0;

            System.Threading.Thread.Sleep(10000);
            foreach (string computer in Globals.onlineComputers)
            {
                try
                {
                    Directory.Delete(@"\\" + computer + @"\C$\Users\Public\Documents\Printer_Migration", true);
                    //Directory.Delete( @"\\" + computer + @"\C$\\Printer_Migration", true );
                }
                catch (Exception ex)
                {
                    errorReport(ex.ToString() + ". Error deleting payload on " + computer);
                }
                finally
                {
                    realProgress += (100f / Globals.onlineComputers.Count);
                    progress = (int)Math.Round(realProgress);
                    worker.ReportProgress(progress, "DELETEPAYLOAD");
                    System.Threading.Thread.Sleep(250);
                }
            }
        }

        private void generateReport(BackgroundWorker worker)
        {
            Globals.table = getData();
            /* int progress = 0;
             double realProgress = 0;
             DataTable build = new DataTable( "Printer" );
             build.Columns.Add( "User" );
             build.Columns.Add( "RetiredPrinters" );
             build.Columns.Add( "DeletionSuccessful" );
             build.Columns.Add( "AddPrinters" );
             build.Columns.Add( "AdditionSuccessful" );
             build.Columns.Add( "Status" );
             build.Columns.Add( "Errors" );
             DataTable table = new DataTable( "Printer" );
             table.Columns.Add( "User" );
             table.Columns.Add( "RetiredPrinters" );
             table.Columns.Add( "DeletionSuccessful" );
             table.Columns.Add( "AddPrinters" );
             table.Columns.Add( "AdditionSuccessful" );
             table.Columns.Add( "Status" );
             table.Columns.Add( "Errors" );
             foreach ( string computer in Globals.onlineComputers )
             {
                 try
                 {
                     DirectoryCopy( @"\\" + computer + @"\C$\Printer_Migration\", useDirectory + @"PrinterLogs\" + computer, true );
                     table.ReadXml( useDirectory + @"PrinterLogs\" + computer + @"\table.xml" );
                     //table.ReadXml( @"S:\TEMP\DeleteIn7Days\PrinterLogs\" + computer + @"\table.xml" );
                     build.Merge( table );
                 }
                 catch ( Exception ex )
                 {
                     errorReport( ex.ToString() + ". Error generating report for " + computer );
                     table.Rows.Add( computer, "---", "---", "---", "---", "ERRORS", "Couldnt retrieve resulting xml file" );
                     build.Merge( table );
                 }
                 finally
                 {
                     realProgress += (100f / Globals.computers.Count);
                     progress = (int)Math.Round( realProgress );
                     worker.ReportProgress( progress, "GENERATEREPORT" );
                     table.Rows.Clear();
                     System.Threading.Thread.Sleep( 1000 );
                 }
             }

             foreach ( string computer in Globals.offlineComputers )
             {
                 table.Rows.Add( computer, "---", "---", "---", "---", "OFFLINE", "Computer is offline" );
                 build.Merge( table );

                 realProgress += (100f / Globals.computers.Count);
                 progress = (int)Math.Round( realProgress );
                 worker.ReportProgress( progress, "GENERATEREPORT" );
                 table.Rows.Clear();
             }

             build.Columns.Add( "On#", typeof( int ) );
             build.Columns.Add( "Off#", typeof( int ) );
             build.Columns.Add( "Error#", typeof( int ) );
             foreach ( DataRow row in build.Rows )
             {
                 if ( row["Status"].Equals( "ONLINE" ) )
                     row["On#"] = 1;
                 if ( row["Status"].Equals( "OFFLINE" ) )
                     row["Off#"] = 1;
                 if ( row["Status"].Equals( "ERRORS" ) )
                     row["Error#"] = 1;
             }

             set.Tables.Add( build );
             Report report = new Report();
             report.ReportTitle = "Printer Migration Report";
             report.ReportSource = set;

             Section section = new Section( "Status", "", Color.Aqua );
             section.IncludeTotal = true;
             report.Sections.Add( section );

             report.TotalFields.Add( "On#" );
             report.TotalFields.Add( "Off#" );
             report.TotalFields.Add( "Error#" );
             report.IncludeTotal = true;

             report.ReportFields.Add( new Field( "User", "User", 12, ALIGN.LEFT ) );
             report.ReportFields.Add( new Field( "RetiredPrinters", "Removed Printers", 12, ALIGN.LEFT ) );
             report.ReportFields.Add( new Field( "DeletionSuccessful", "Removal Successful", 12, ALIGN.LEFT ) );
             report.ReportFields.Add( new Field( "AddPrinters", "Added Printers", 12, ALIGN.LEFT ) );
             report.ReportFields.Add( new Field( "AdditionSuccessful", "Addition Successful", 12, ALIGN.LEFT ) );
             report.ReportFields.Add( new Field( "Errors", "Errors", 12, ALIGN.LEFT ) );
             report.ReportFields.Add( new Field( "On#", "On#", 1, ALIGN.LEFT ) );
             report.ReportFields.Add( new Field( "Off#", "Off#", 1, ALIGN.LEFT ) );
             report.ReportFields.Add( new Field( "Error#", "Error#", 1, ALIGN.LEFT ) );

             report.SaveReport( useDirectory + "Report.htm" );*/
        }

        private DataTable getData()
        {
            DataTable temp = new DataTable("Migration Report");
            using (SqlConnection connection = new SqlConnection(Globals.dbConnection))
            {
                connection.Open();
                try
                {
                    using (SqlCommand command = new SqlCommand("SELECT * FROM Main", connection))
                    {
                        temp.Load(command.ExecuteReader());
                    }
                }
                catch (Exception ex)
                {
                    errorReport(ex.ToString());
                }
            }

            return temp;
        }

        private void executePayload(BackgroundWorker worker)
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.System) + "\\owexec.exe";
            int progress = 0;
            double realProgress = 0;

            if (!File.Exists(path))
            {
                System.Reflection.Assembly myAssembly = System.Reflection.Assembly.GetExecutingAssembly();
                Stream myStream = myAssembly.GetManifestResourceStream("Printer_Migration.owexec.exe");

                using (FileStream fileStream = File.Create(path))
                {
                    myStream.CopyTo(fileStream);
                }
            }

            System.Diagnostics.Process proc = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            foreach (string computer in Globals.onlineComputers)
            {
                try
                {
                    startInfo.FileName = path;
                    startInfo.Arguments = "-nowait -c " + computer + " -k " + useDirectory + @"run.bat -p ""-a " +
                                                            tbAddPrinters.Text + " -d " + tbDeletePrinters.Text + @"""";
                    startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                    proc.StartInfo = startInfo;
                    proc.Start();
                    System.Threading.Thread.Sleep(900);
                }
                catch (Exception ex)
                {
                    errorReport(ex.ToString() + ". Error executing CalRecycle_Printer_Upgrade on " + computer);
                }
                finally
                {
                    realProgress += ((100f / Globals.onlineComputers.Count)) / 2;
                    progress = (int)Math.Round(realProgress);
                    worker.ReportProgress(progress, "EXECUTEPAYLOAD");
                }
            }

            // kill some time to let computers finish, why I divide by 2 on real progress
            foreach (string waitComputer in Globals.onlineComputers)
            {
                realProgress += ((100f / Globals.onlineComputers.Count)) / 2;
                progress = (int)Math.Round(realProgress);
                worker.ReportProgress(progress, "EXECUTEPAYLOAD");
                System.Threading.Thread.Sleep(2000);
            }
        }

        private void transferPayload(BackgroundWorker worker)
        {
            System.Reflection.Assembly myAssembly = System.Reflection.Assembly.GetExecutingAssembly();
            Stream myStream;
            int progress = 0;
            double realProgress = 0;
            List<String> remove = new List<String>();

            foreach (string computer in Globals.onlineComputers)
            {
                try
                {
                    if (!Directory.Exists(@"\\" + computer + @"\C$\Users\Public\Documents\Printer_Migration"))
                    {
                        Directory.CreateDirectory(@"\\" + computer + @"\C$\Users\Public\Documents\Printer_Migration");
                    }
                    myStream = myAssembly.GetManifestResourceStream("Printer_Migration.Payload.CalRecycle_Printer_Upgrade.exe");
                    using (var fileStream = File.Create(@"\\" + computer + @"\C$\Users\Public\Documents\Printer_Migration\CalRecycle_Printer_Upgrade.exe"))
                    {
                        myStream.CopyTo(fileStream);
                    }

                    myStream = myAssembly.GetManifestResourceStream("Printer_Migration.Payload.InstallDriver.bat");
                    using (var fileStream = File.Create(@"\\" + computer + @"\C$\Users\Public\Documents\Printer_Migration\InstallDriver.bat"))
                    {
                        myStream.CopyTo(fileStream);
                    }

                    myStream = myAssembly.GetManifestResourceStream("Printer_Migration.Payload.SharpCert.cer");
                    using (var fileStream = File.Create(@"\\" + computer + @"\C$\Users\Public\Documents\Printer_Migration\SharpCert.cer"))
                    {
                        myStream.CopyTo(fileStream);
                    }

                    myStream = myAssembly.GetManifestResourceStream("Printer_Migration.Payload.run.bat");
                    using (var fileStream = File.Create(@"\\" + computer + @"\C$\Users\Public\Documents\Printer_Migration\run.bat"))
                    {
                        myStream.CopyTo(fileStream);
                    }
                }
                catch (Exception ex)
                {
                    errorReport(ex.ToString() + ". Error copying CalRecycle_Printer_Upgrade to " + computer);
                    remove.Add(computer);
                }
                finally
                {
                    realProgress += (100f / Globals.onlineComputers.Count);
                    progress = (int)Math.Round(realProgress);
                    worker.ReportProgress(progress, "TRANSFERPAYLOAD");
                }
            }

            if (remove.Count > 0)
            {
                foreach (string computerRemove in remove)
                {
                    Globals.onlineComputers.Remove(computerRemove);
                    Globals.offlineComputers.Add(computerRemove);
                }
            }
        }

        private void pingComputers(BackgroundWorker worker)
        {
            Ping ping = new Ping();
            PingReply reply;
            int progress = 0;
            double realProgress = 0;

            foreach (string computer in Globals.computers)
            {
                try
                {
                    reply = ping.Send(computer);
                    if (reply.Status == System.Net.NetworkInformation.IPStatus.Success)
                    {
                        Globals.onlineComputers.Add(computer);
                    }
                    else
                    {
                        Globals.offlineComputers.Add(computer);
                    }
                }
                catch (System.Net.NetworkInformation.PingException)
                {
                    Globals.offlineComputers.Add(computer);
                }
                finally
                {
                    realProgress += (100f / Globals.computers.Count);
                    progress = (int)Math.Round(realProgress);
                    worker.ReportProgress(progress, "PINGCOMPUTERS");
                }
            }
        }

        private void generateComputerNames(BackgroundWorker worker)
        {
            string computerName, fInitial, LName, OS = "W8-";
            int LNameSize;
            int progress = 0;
            double realProgress = 0;

            foreach (string user in Globals.users)
            {
                try
                {
                    fInitial = user.ElementAt(0).ToString();
                    LNameSize = user.IndexOf('@') - user.IndexOf('.') - 1;

                    if (LNameSize <= 7)
                        LName = user.Substring(user.IndexOf('.') + 1, LNameSize);
                    else
                        LName = user.Substring(user.IndexOf('.') + 1, 7);

                    computerName = OS + fInitial + LName;
                    Globals.computers.Add(computerName);
                }
                catch (Exception ex)
                {
                    errorReport(ex.ToString());
                    Globals.offlineComputers.Add(user);
                }
                finally
                {
                    realProgress += (100f / Globals.users.Count);
                    progress = (int)Math.Round(realProgress);
                    worker.ReportProgress(progress, "GENERATECOMPUTERNAMES");
                }
            }

            worker.ReportProgress(100, "GENERATECOMPUTERNAMES");
        }

        #endregion bwRun

        #region Logs

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

        private void getUsers()
        {
            Stream myStream = null;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if ((myStream = openFileDialog1.OpenFile()) != null)
                    {
                        int cur;
                        string btyesToString = "", modifiedBytesToString = "";
                        using (myStream)
                        {
                            while ((cur = myStream.ReadByte()) != -1)
                            {
                                if (cur != 44)
                                {
                                    Globals.temp.Add((byte)cur);
                                }
                                else
                                {
                                    btyesToString = System.Text.Encoding.Default.GetString(Globals.temp.ToArray());
                                    modifiedBytesToString = btyesToString.Replace(System.Environment.NewLine, "");
                                    Globals.users.Add(modifiedBytesToString);
                                    Globals.temp.Clear();
                                }
                            }

                            // Get last user if any
                            btyesToString = System.Text.Encoding.Default.GetString(Globals.temp.ToArray());
                            modifiedBytesToString = btyesToString.Replace(System.Environment.NewLine, "");
                            Globals.users.Add(modifiedBytesToString);

                            Globals.temp.Clear();
                            Globals.users.RemoveAll(string.IsNullOrWhiteSpace);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }
    }
}