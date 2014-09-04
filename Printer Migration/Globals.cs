using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace Printer_Migration
{
    public class Globals
    {
        public static List<byte> temp = new List<byte>();
        public static List<String> users = new List<String>();
        public static List<String> computers = new List<String>();
        public static List<String> onlineComputers = new List<String>();
        public static List<String> offlineComputers = new List<String>();
        public static List<String> printersOnServer = new List<String>();
        public static DataTable table;
        public static string userXLSPath;
        public static string dbConnection = @"Data Source=W8-RKOEN;Initial Catalog=PrinterMigration;User ID=sa;Password=p@ssw0rd";
    }
}