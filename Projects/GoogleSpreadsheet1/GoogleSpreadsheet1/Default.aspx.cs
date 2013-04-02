using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;
using Google.GData.Client;
using Google.GData.Extensions;
using Google.GData.Spreadsheets;

namespace GoogleSpreadsheet1
{
    public partial class _Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
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
