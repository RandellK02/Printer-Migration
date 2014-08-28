using HTMLReportEngine;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;

namespace Test
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            string path = @"C:\Users\rkoen\";
            string[] users = { "Randy", "Jacky", "Philip", "Reese", "Nick" };
            string retiredPrinter = "08-colormfp-03";
            bool successfulDelete = true, sAdd = false, offline = false, online = true;

            DataSet ds = new DataSet( "Summary" );

            DataTable dtUsers = new DataTable();
            dtUsers.Columns.Add( "User", typeof( string ) );
            dtUsers.Columns.Add( "retiredPrinter", typeof( string ) );
            dtUsers.Columns.Add( "successfulDelete", typeof( bool ) );
            dtUsers.Columns.Add( "sAdd", typeof( bool ) );
            dtUsers.Columns.Add( "offline", typeof( bool ) );
            dtUsers.Columns.Add( "Online", typeof( bool ) );
            dtUsers.Columns.Add( "project", typeof( int ) );

            foreach ( string user in users )
            {
                dtUsers.Rows.Add( user, retiredPrinter, successfulDelete, sAdd, offline, online, 1 );
            }

            ds.Tables.Add( dtUsers );
            ds.WriteXml( path + "source.xml" );

            DataSet copy = new DataSet( "copy" );
            copy.ReadXml( path + "source.xml" );
            Report report = new Report();
            report.ReportTitle = "Printer Migration Report";
            report.ReportSource = copy;

            report.IncludeTotal = true;
            report.TotalFields.Add( "project" );

            Section release = new Section( "Online", "Online: ", Color.Aqua );
            release.IncludeTotal = true;
            Section id = new Section( "User", "User: " );
            id.IncludeTotal = true;
            release.SubSection = id;
            report.Sections.Add( release );
            report.ReportFields.Add( new Field( "project", "ID", 1, ALIGN.LEFT ) );
            report.ReportFields.Add( new Field( "retiredPrinter", "Retired Printer", 12, ALIGN.LEFT ) );
            report.ReportFields.Add( new Field( "successfulDelete", "Deletion Successful", 12, ALIGN.LEFT ) );
            report.ReportFields.Add( new Field( "sAdd", "Add Successful", 12, ALIGN.LEFT ) );

            try
            {
                report.SaveReport( path + "Report.htm" );
            }
            catch ( Exception ex )
            {
                Console.WriteLine( ex.ToString() );
                Console.ReadKey();
            }
        }
    }
}