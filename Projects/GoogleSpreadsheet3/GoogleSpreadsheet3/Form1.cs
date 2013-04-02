using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Google.GData.Client;
using Google.GData.Extensions;
using Google.GData.Spreadsheets;


namespace GoogleSpreadsheet3
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            SpreadsheetsService myService = new SpreadsheetsService("exampleApp-2");
            myService.setUserCredentials("mike3889@gmail.com", "christmas.25");
            SpreadsheetQuery query = new SpreadsheetQuery();
            SpreadsheetFeed feed = myService.Query(query);

            Console.WriteLine("Your spreadsheets:");
            foreach (SpreadsheetEntry entry in feed.Entries)
            {
                Console.WriteLine(entry.Title.Text);
            }
        }
    }
}
