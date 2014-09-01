using System;
using System.Collections;
using System.Collections.Generic;
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
        public static string userXLSPath;
    }
}