using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SqlConnection myConnection = new SqlConnection("server=localhost;" +
                          "Trusted_Connection=yes;" +
                          "database=quickbase;" +
                          "connection timeout=30");
            SqlConnection dataConnection = new SqlConnection("server=localhost;" +
                                       "Trusted_Connection=yes;" +
                                       "database=quickbase;" +
                                       "connection timeout=30");
            SqlCommand mySqlCommand = myConnection.CreateCommand();
            SqlCommand dataSqlCommand = dataConnection.CreateCommand();
            myConnection.Open();
            dataConnection.Open();
            SqlCommand myCommand = new SqlCommand("select * from billable_leases", myConnection);


            SqlDataReader tenantsSqlDataReader = myCommand.ExecuteReader();

            int TenantNameColPos = tenantsSqlDataReader.GetOrdinal("Tenant Name");
            int BillingBaseRentMonthlyColPos = tenantsSqlDataReader.GetOrdinal("BillingBaseRentMonthly");
            int BillableColPos = tenantsSqlDataReader.GetOrdinal("Billable");
            int BillingOperatingExpenseMonthlyColPos = tenantsSqlDataReader.GetOrdinal("BillingOperatingExpenseMonthly");
            int BillingTaxesMonthlyColPos = tenantsSqlDataReader.GetOrdinal("BillingTaxesMonthly");
            int BillingMonthColPos = tenantsSqlDataReader.GetOrdinal("Billing Month");
            int YearColPos = tenantsSqlDataReader.GetOrdinal("Year");
            int TenantReferenceKeyColPos = tenantsSqlDataReader.GetOrdinal("TenantReferenceKey");

            while (tenantsSqlDataReader.Read())
            {
                string TenantName = tenantsSqlDataReader.GetString(TenantNameColPos);
                bool Billable = tenantsSqlDataReader.GetBoolean(BillableColPos);
                decimal BillingBaseRentMonthly = tenantsSqlDataReader.GetDecimal(BillingBaseRentMonthlyColPos);
                decimal BillingTaxesMonthly = tenantsSqlDataReader.GetDecimal(BillingTaxesMonthlyColPos);
                decimal BillingOperatingExpenseMonthly = tenantsSqlDataReader.GetDecimal(BillingOperatingExpenseMonthlyColPos);
                string BillingMonth = tenantsSqlDataReader.GetString(BillingMonthColPos);
                string BillingYear = tenantsSqlDataReader.GetString(YearColPos);
                string TenantReferenceKey = tenantsSqlDataReader.GetString(TenantReferenceKeyColPos);


                //               if (tenantsSqlDataReader.GetDecimal(BillingBaseRentMonthlyColPos)!=null) 
                //{
                //    decimal BillingBaseRentMonthly = tenantsSqlDataReader.GetDecimal(BillingBaseRentMonthlyColPos);    
                //}
                //else
                //{
                //    decimal BillingBaseRentMonthly = 0; 

                //}

                SqlCommand dataCommand = new SqlCommand("INSERT InvoiceDetails (TenantName, Billable, Amount, BillingYear, BillingMonth, TenantReferenceKey, Type) VALUES ('" + TenantName + "', 1," + BillingBaseRentMonthly + ",'" + BillingYear + "','" + BillingMonth + "','" + TenantReferenceKey + "','Rent')", dataConnection);

                SqlCommand data1Command = new SqlCommand("INSERT InvoiceDetails (TenantName, Billable, Amount, BillingYear, BillingMonth, TenantReferenceKey, Type) VALUES ('" + TenantName + "', 1," + BillingOperatingExpenseMonthly + ",'" + BillingYear + "','" + BillingMonth + "','" + TenantReferenceKey + "','OpExp')", dataConnection);
                SqlCommand data2Command = new SqlCommand("INSERT InvoiceDetails (TenantName, Billable, Amount, BillingYear, BillingMonth, TenantReferenceKey, Type) VALUES ('" + TenantName + "', 1," + BillingTaxesMonthly + ",'" + BillingYear + "','" + BillingMonth + "','" + TenantReferenceKey + "','Taxes')", dataConnection);
                dataCommand.ExecuteNonQuery();
                data1Command.ExecuteNonQuery();
                data2Command.ExecuteNonQuery();
            }
            tenantsSqlDataReader.Close();

            // Process Hodges Room charges
            DateTime LastMonthBillingDate = new DateTime(2009, 8, 20);
            DateTime CurrentMonthBillingDate = new DateTime(2009, 9, 20);

            SqlCommand data3Command = new SqlCommand("INSERT INTO InvoiceDetails (TenantName, Amount, BillingYear, BillingMonth, BillingDate, Type) SELECT [Related Tenant], [Total Revenue], Year, Month, [Event Date], 'Hodges Room' WHERE [Event Date] > " + LastMonthBillingDate + "AND [Event Date] < " + CurrentMonthBillingDate + ")", dataConnection);
            data3Command.ExecuteNonQuery();

            myConnection.Close();

        }
    }
}
