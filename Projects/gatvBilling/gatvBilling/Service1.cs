using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Data.SqlClient;
using System.Data.SqlTypes;



namespace gatvBilling
{
    public partial class Service1 : ServiceBase
    {
        SqlConnection myConnection = new SqlConnection("server=localhost;" +
                                       "Trusted_Connection=yes;" +
                                       "database=quickbase;" +
                                       "connection timeout=30");
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            
        }

        
        protected override void OnStop()
        {
        }
    public string ReadSomeData()
    {
            try
    {
        myConnection.Open();
        DataTable tblDatabases = myConnection.GetSchema ("Databases");
   myConnection.Close();
 
   foreach (DataRow row in tblDatabases.Rows)
   {
      Console.WriteLine ("Database: " + row["database_name"]);
   }
    }
    catch(Exception e)
    {
        Console.WriteLine(e.ToString());
    }

            return "a";
    }
     
    
    
    }


}
