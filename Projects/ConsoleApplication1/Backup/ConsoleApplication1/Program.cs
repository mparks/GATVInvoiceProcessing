using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;


namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
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

            //        SqlCommand myInvoiceHeaderCommand = new SqlCommand("select * from BillableTenants", myConnection);
            SqlCommand myDeleteCommand1 = new SqlCommand("Delete  from InvoiceHeaders", myConnection);
            myDeleteCommand1.ExecuteNonQuery();
            SqlCommand myInvoiceHeaderCommand = new SqlCommand("INSERT INTO InvoiceHeaders (TenantName, BillingMonth, BillingYear, TenantReferenceKey, PrimarySuite) (Select [Tenant Name], BillingMonth, BillingYear, [Tenant Name] + BillingMonth + BillingYear, [Primary Suite] FROM Billable_Leases)", dataConnection);
            myInvoiceHeaderCommand.ExecuteNonQuery();
            SqlCommand myInvoiceHeaderCommand1 = new SqlCommand("INSERT INTO InvoiceHeaders (TenantName, BillingMonth, BillingYear, TenantReferenceKey, PrimarySuite) (Select [Tenant Name], BillingMonth, BillingYear, [Tenant Name] + BillingMonth + BillingYear, '000' FROM Billable_Tenants)", dataConnection);
            myInvoiceHeaderCommand1.ExecuteNonQuery();


            // Process lease suite details cost data and insert records into InvoiceDetails table.            
            SqlCommand myCommand = new SqlCommand("select * from LeaseSuiteDetailsBillingData", myConnection);

            SqlCommand myDeleteCommand = new SqlCommand("Delete  from InvoiceDetails", myConnection);
            myDeleteCommand.ExecuteNonQuery();


            SqlDataReader tenantsSqlDataReader = myCommand.ExecuteReader();

            int TenantNameColPos = tenantsSqlDataReader.GetOrdinal("Tenant Name");
            int BillingBaseRentMonthlyColPos = tenantsSqlDataReader.GetOrdinal("Base Monthly Rental Amount");
            //       int BillableColPos = tenantsSqlDataReader.GetOrdinal("Billable");
            int BillingOperatingExpenseMonthlyColPos = tenantsSqlDataReader.GetOrdinal("MonthlyOperatingExpense");
            int BillingTaxesMonthlyColPos = tenantsSqlDataReader.GetOrdinal("MonthlyTaxesExpense");
            int BillingMonthColPos = tenantsSqlDataReader.GetOrdinal("BillingMonth");
            int YearColPos = tenantsSqlDataReader.GetOrdinal("BillingYear");
            int TenantReferenceKeyColPos = tenantsSqlDataReader.GetOrdinal("TenantReferenceKey");
            int SuiteColPos = tenantsSqlDataReader.GetOrdinal("Suite");
            int BillingDateColPos = tenantsSqlDataReader.GetOrdinal("Billing Date");
            int BillingWallExpenseColPos = tenantsSqlDataReader.GetOrdinal("MonthlyWallAmount1");
            int BillingOfficeExpenseColPos = tenantsSqlDataReader.GetOrdinal("MonthlyOfficeAmount1");
            int BillingCubeExpenseColPos = tenantsSqlDataReader.GetOrdinal("MonthlyCubeAmount1");
            int BillingFurnitureExpenseColPos = tenantsSqlDataReader.GetOrdinal("MonthlyFurnitureAmount1");


            while (tenantsSqlDataReader.Read())
            {
                string TenantName = tenantsSqlDataReader.GetString(TenantNameColPos);
                //                bool Billable = tenantsSqlDataReader.GetBoolean(BillableColPos);
                decimal BillingBaseRentMonthly = tenantsSqlDataReader.GetDecimal(BillingBaseRentMonthlyColPos);
                decimal BillingTaxesMonthly = tenantsSqlDataReader.GetDecimal(BillingTaxesMonthlyColPos);
                decimal BillingOperatingExpenseMonthly = tenantsSqlDataReader.GetDecimal(BillingOperatingExpenseMonthlyColPos);
                int BillingMonth = tenantsSqlDataReader.GetInt32(BillingMonthColPos);
                int BillingYear = tenantsSqlDataReader.GetInt32(YearColPos);
                string TenantReferenceKey = tenantsSqlDataReader.GetString(TenantReferenceKeyColPos);
                string Suite = tenantsSqlDataReader.GetString(SuiteColPos);
                DateTime BillingDate = tenantsSqlDataReader.GetDateTime(BillingDateColPos);

                decimal BillingWallMonthly = tenantsSqlDataReader.GetDecimal(BillingWallExpenseColPos);
                decimal BillingOfficeMonthly = tenantsSqlDataReader.GetDecimal(BillingOfficeExpenseColPos);
                decimal BillingCubeMonthly = tenantsSqlDataReader.GetDecimal(BillingCubeExpenseColPos);
                decimal BillingFurnitureMonthly = tenantsSqlDataReader.GetDecimal(BillingFurnitureExpenseColPos);

                //               if (tenantsSqlDataReader.GetDecimal(BillingBaseRentMonthlyColPos)!=null) 
                //{
                //    decimal BillingBaseRentMonthly = tenantsSqlDataReader.GetDecimal(BillingBaseRentMonthlyColPos);    
                //}
                //else
                //{
                //    decimal BillingBaseRentMonthly = 0; 

                //}

                string strDateTime = BillingMonth + "/1/" + BillingYear;
                DateTime myDateTime = Convert.ToDateTime(strDateTime);

                SqlCommand dataCommand = new SqlCommand("INSERT InvoiceDetails (TenantName, Billable, Amount, BillingYear, BillingMonth, TenantReferenceKey, Type, Description, BillingDate) VALUES ('" + TenantName + "', 1," + BillingBaseRentMonthly + ",'" + BillingYear + "','" + BillingMonth + "','" + TenantReferenceKey + "','Rent'," + "'Rent for Suite " + Suite + "', '" + BillingDate + "')", dataConnection);

                SqlCommand data1Command = new SqlCommand("INSERT InvoiceDetails (TenantName, Billable, Amount, BillingYear, BillingMonth, TenantReferenceKey, Type, Description, BillingDate) VALUES ('" + TenantName + "', 1," + BillingOperatingExpenseMonthly + ",'" + BillingYear + "','" + BillingMonth + "','" + TenantReferenceKey + "','OpExp'," + "'Operating Expense for Suite " + Suite + "', '" + BillingDate + "')", dataConnection);
                SqlCommand data2Command = new SqlCommand("INSERT InvoiceDetails (TenantName, Billable, Amount, BillingYear, BillingMonth, TenantReferenceKey, Type, Description, BillingDate) VALUES ('" + TenantName + "', 1," + BillingTaxesMonthly + ",'" + BillingYear + "','" + BillingMonth + "','" + TenantReferenceKey + "','Taxes'," + "'Taxes for Suite " + Suite + "', '" + BillingDate + "')", dataConnection);
                if (BillingWallMonthly > 0)
                {
                    SqlCommand data3Command = new SqlCommand("INSERT InvoiceDetails (TenantName, Billable, Amount, BillingYear, BillingMonth, TenantReferenceKey, Type, Description, BillingDate) VALUES ('" + TenantName + "', 1," + BillingWallMonthly + ",'" + BillingYear + "','" + BillingMonth + "','" + TenantReferenceKey + "','Furniture'," + "'Wall Charge for Suite " + Suite + "', '" + BillingDate + "')", dataConnection);
                    data3Command.ExecuteNonQuery();
                }
                if (BillingOfficeMonthly > 0)
                {
                    SqlCommand data4Command = new SqlCommand("INSERT InvoiceDetails (TenantName, Billable, Amount, BillingYear, BillingMonth, TenantReferenceKey, Type, Description, BillingDate) VALUES ('" + TenantName + "', 1," + BillingOfficeMonthly + ",'" + BillingYear + "','" + BillingMonth + "','" + TenantReferenceKey + "','Furniture'," + "'Office Charge for Suite " + Suite + "', '" + BillingDate + "')", dataConnection);
                    data4Command.ExecuteNonQuery();
                }
                if (BillingCubeMonthly > 0)
                {
                    SqlCommand data5Command = new SqlCommand("INSERT InvoiceDetails (TenantName, Billable, Amount, BillingYear, BillingMonth, TenantReferenceKey, Type, Description, BillingDate) VALUES ('" + TenantName + "', 1," + BillingCubeMonthly + ",'" + BillingYear + "','" + BillingMonth + "','" + TenantReferenceKey + "','Furniture'," + "'Cube Charge for Suite " + Suite + "', '" + BillingDate + "')", dataConnection);
                    data5Command.ExecuteNonQuery();
                }
                if (BillingFurnitureMonthly > 0)
                {
                    SqlCommand data6Command = new SqlCommand("INSERT InvoiceDetails (TenantName, Billable, Amount, BillingYear, BillingMonth, TenantReferenceKey, Type, Description, BillingDate) VALUES ('" + TenantName + "', 1," + BillingFurnitureMonthly + ",'" + BillingYear + "','" + BillingMonth + "','" + TenantReferenceKey + "','Furniture'," + "'Furniture Charge for Suite " + Suite + "', '" + BillingDate + "')", dataConnection);
                    data6Command.ExecuteNonQuery();
                }

                dataCommand.ExecuteNonQuery();
                data1Command.ExecuteNonQuery();
                data2Command.ExecuteNonQuery();
            }
            tenantsSqlDataReader.Close();

            // Process Hodges Room charges
    //        DateTime LastMonthBillingDate = new DateTime(2009, 7, 20);
    //        DateTime CurrentMonthBillingDate = new DateTime(2009, 8, 20);


    //        SqlCommand data7Command = new SqlCommand("INSERT INTO InvoiceDetails (TenantName, Amount, BillingMonth, BillingYear, BillingDate, TenantReferenceKey, Type, Description) (Select [Related Tenant], TotalRevenue, BillingMonth, BillingYear,[Event Date],[Related Tenant]+ BillingMonth + BillingYear, 'Hodges Room', desc1 FROM HodgesRoomBilling)", dataConnection);

    //        data7Command.ExecuteNonQuery();


            // Process Credits and Charges

            SqlCommand data8Command = new SqlCommand("INSERT INTO InvoiceDetails (TenantName, Amount, BillingMonth, BillingYear, TenantReferenceKey, Type, Description, BillingDate) (Select [Tenant Name], Amount, BillingMonth, BillingYear,[Tenant Name]+ BillingMonth + BillingYear, Type, Description, [Billing Date] FROM CreditsAndChargesBilling)", dataConnection);
            data8Command.ExecuteNonQuery();




            // Process Parkers


            SqlCommand data9Command = new SqlCommand("INSERT INTO InvoiceDetails (TenantName, Amount, BillingMonth, BillingYear, TenantReferenceKey, Type, Description, BillingDate) (Select TenantName, stdMonthlyParking, BillingMonth, BillingYear,TenantName+ BillingMonth + BillingYear, 'Parking', desc1, [Billing Date] FROM standardparking WHERE stdMonthlyParking>0)", dataConnection);
            data9Command.ExecuteNonQuery();

            SqlCommand data10Command = new SqlCommand("INSERT INTO InvoiceDetails (TenantName, Amount, BillingMonth, BillingYear, TenantReferenceKey, Type, Description, BillingDate) (Select TenantName, OverAllocationMonthlyParking, BillingMonth, BillingYear,TenantName+ BillingMonth + BillingYear, 'Parking', desc1, [Billing Date] FROM OverAllotmentParking WHERE OverAllocationMonthlyParking>0)", dataConnection);
            data10Command.ExecuteNonQuery();

            SqlCommand data18Command = new SqlCommand("INSERT InvoiceDetails (TenantName, Billable, Amount, BillingYear, BillingMonth, TenantReferenceKey, Type, Description) VALUES ('Midtown Park Associates, LLC (Georgia Power)', 1," + 1033.40 + ",'2010','01','Midtown Park Associates, LLC (Georgia Power)012010','Parking','Parking charge for 20 parkers')", dataConnection);
            data18Command.ExecuteNonQuery();

            // Process Payments
            //           SqlCommand data17Command = new SqlCommand("INSERT INTO InvoiceDetails (TenantName, Amount, BillingMonth, BillingYear, TenantReferenceKey, Type, Description) (Select [Related Tenant], Payment, BillingMonth, BillingYear,TenantName+ BillingMonth + BillingYear, 'Payment', 'Payment' FROM Payments)", dataConnection);
            SqlCommand data17Command = new SqlCommand("INSERT INTO InvoiceDetails (TenantName, Amount, BillingMonth, BillingYear, TenantReferenceKey, Type, Description, BillingDate) (Select [related tenant], Payment, BillingMonth, BillingYear, [related tenant] + BillingMonth + BillingYear, 'Payment', 'Payment', BillingDate FROM Payments)", dataConnection);

            data17Command.ExecuteNonQuery();

            // Create Beginning Balances


  //          label1.Text = "Complete";
            myConnection.Close();
 
   
   
    }
 

        }
    }

