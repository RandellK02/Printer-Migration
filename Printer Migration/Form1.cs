using System;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Windows.Forms;

namespace Printer_Migration
{
    public partial class Form1 : Form
    {
        public Stream myStream;
        private static DataTable summary;
        private static string useDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments) + @"Printer_Migration\";
        private static string logFile = useDirectory + "printer.log";
        private static string errorLogFile = useDirectory + "error.log";

        public Form1()
        {
            InitializeComponent();
            myStream = null;
            summary = new DataTable();
            summary.Columns.Add("User", typeof(string));
            /* summary.Columns.Add("Retired Printer", typeof(bool));
             summary.Columns.Add("Retired Removed", typeof(bool));
             summary.Columns.Add("New Printer Installed", typeof(bool));*/

            if (!Directory.Exists(useDirectory))
            {
                Directory.CreateDirectory(useDirectory);
            }
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if ((myStream = openFileDialog1.OpenFile()) != null)
                    {
                        if (bwUserImport.IsBusy != true)
                        {
                            bwUserImport.RunWorkerAsync();
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }

            //getUsers();
            //generateComputerNames();
        }

        #region Import

        private void bwUserImport_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            Object thisLock = new Object();
            int myStreamLength = (int)myStream.Length;
            int cur;
            int progress = 0;
            double realProgress = 0;
            string btyesToString = "", modifiedBytesToString = "";

            lock (thisLock)
            {
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

                        realProgress += (100f / myStreamLength);
                        progress = (int)Math.Round(realProgress);
                        worker.ReportProgress(progress);
                    }

                    // Get last user if any
                    btyesToString = System.Text.Encoding.Default.GetString(Globals.temp.ToArray());
                    modifiedBytesToString = btyesToString.Replace(System.Environment.NewLine, "");
                    Globals.users.Add(modifiedBytesToString);

                    Globals.temp.Clear();
                    Globals.users.RemoveAll(string.IsNullOrWhiteSpace);
                    worker.ReportProgress(100);
                }
            }
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
            }
        }

        private void bwRun_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            generateComputerNames(worker);
            pingComputers(worker);
            transferPayload(worker);
            executePayload(worker);
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

            foreach (string computer in Globals.onlineComputers)
            {
                try
                {
                    System.Diagnostics.Process.Start(path, "-nowait -c " + computer + " -k " + useDirectory + @"Payload.exe -p ""-a " +
                                                            tbAddPrinters.Text + " -d " + tbDeletePrinters.Text + @"""");
                    System.Threading.Thread.Sleep(100);
                }
                catch (Exception ex)
                {
                    errorReport(ex.ToString() + ". Error executing payload on " + computer);
                }
                finally
                {
                    realProgress += (100f / Globals.onlineComputers.Count);
                    progress = (int)Math.Round(realProgress);
                    worker.ReportProgress(progress, "EXECUTEPAYLOAD");
                }
            }
        }

        private void transferPayload(BackgroundWorker worker)
        {
            System.Reflection.Assembly myAssembly = System.Reflection.Assembly.GetExecutingAssembly();
            Stream myStream = myAssembly.GetManifestResourceStream("Printer_Migration.Payload.exe");
            int progress = 0;
            double realProgress = 0;
            Globals.onlineComputers.Add("Test");
            foreach (string computer in Globals.onlineComputers)
            {
                try
                {
                    if (!Directory.Exists(@"\\" + computer + @"\C$\Users\Public\Documents\Printer_Migration"))
                    {
                        Directory.CreateDirectory(@"\\" + computer + @"\C$\Users\Public\Documents\Printer_Migration");
                    }
                    using (var fileStream = File.Create(@"\\" + computer + @"\C$\Users\Public\Documents\Printer_Migration\Payload.exe"))
                    {
                        myStream.CopyTo(fileStream);
                    }
                }
                catch (Exception ex)
                {
                    errorReport(ex.ToString() + ". Error copying payload to " + computer);
                    Globals.onlineComputers.Remove(computer);
                    Globals.offlineComputers.Add(computer);
                }
                finally
                {
                    realProgress += (100f / Globals.onlineComputers.Count);
                    progress = (int)Math.Round(realProgress);
                    worker.ReportProgress(progress, "TRANSFERPAYLOAD");
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
                fInitial = user.ElementAt(0).ToString();
                LNameSize = user.IndexOf('@') - user.IndexOf('.') - 1;

                if (LNameSize <= 7)
                    LName = user.Substring(user.IndexOf('.') + 1, LNameSize);
                else
                    LName = user.Substring(user.IndexOf('.') + 1, 7);

                computerName = OS + fInitial + LName;
                Globals.computers.Add(computerName);

                realProgress += (100f / Globals.users.Count);
                progress = (int)Math.Round(realProgress);
                worker.ReportProgress(progress, "GENERATECOMPUTERNAMES");
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