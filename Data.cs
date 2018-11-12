////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// Alivia Houdek 
// 08.20.2018 
// Universal Returns Automation 
// Run on every Saturday AND every first of the month

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Data.Odbc;
using System.Data;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Net.Mail;
using System.Net;
using System.Collections;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace AHK_Universal_Returns
{
    class Data
    {
        // Change this variable temporarily to change the date the program believes today is (if the program failed to run). Don't forget to reset after!

        static int dateAdjust = 3; // change date manually if re-running failed program (How many days ago is the day it was supposed to run on?)
        bool validator = true; // turns on data validation for dates if true; change to false if re-running failed program fails because of validation

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        // Dates
        static int leap;
        static DateTime date = DateTime.Now;
        static DateTime firstDate = new DateTime(date.Year, date.Month, 1);
        static DateTime secondDate = firstDate.AddMonths(1).AddDays(-1);
        //static string firstMonthDate = new DateTime(date.Year, date.Month, 1).ToString("yyyy-MM-dd");
        //static string secondMonthDate = firstDate.AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd");
        string newDate = date.ToString("MM/dd/yyyy"); // update current date format for mySQL Update
        DateTime curDt = date.AddDays(-dateAdjust); // current date or specified date of program being run

        // Connections
        string selectSQL = "";
        string bhConnStr = "";
        string sqlConnStr = "Driver={MYSQL ODBC 5.1 Driver};Server=amc-mysql;uid=univR;pwd=RXCKN1Vum4MpclJH;db=UniversalReturns;"; //RXCKN1Vum4MpclJH just in case

        // Client/Creditor Group Object Stuff
        ReturnAccount currentAccountObj;
        ClientGroup creditorGroup;

        // Data Tables
        System.Data.DataTable bhDt = new System.Data.DataTable(); // BH data for debtors
        System.Data.DataTable sqDt = new System.Data.DataTable(); // Data from original macro for clients/client groups
        System.Data.DataTable bhCreditors = new System.Data.DataTable(); // stores creditor numbers from BH for validation
        System.Data.DataTable validStatuses = new System.Data.DataTable(); // for bh status validation
        //System.Data.DataTable bankos = new System.Data.DataTable();
        System.Data.DataTable settleds = new System.Data.DataTable();
        System.Data.DataTable returns = new System.Data.DataTable();
        System.Data.DataTable customTable = new System.Data.DataTable();

        // names of all banko-specific columns
        string[] bankoColNames = { "Chpt / Case Number", "Filing Date", "Discharge Date", "Proof Date", "Dismissal Date" };

        // Excel Output Variables
        object misValue = System.Reflection.Missing.Value;
        string filePath; // says it isn't used but it is
        List<string> emailList = new List<string>(); // list of all email addresses in DB for client
        List<System.Data.DataTable> splitReturnTablesList = new List<System.Data.DataTable>();
        List<String> bankoStatusCodes = new List<String>(); // store status codes from Bloodhound that are for bankruptcies
        List<string> bankoTableNames = new List<string>();

        // Dictionaries
        Dictionary<string, ClientGroup> creditorGroupDict = new Dictionary<string, ClientGroup>();
        Dictionary<int, ReturnAccount> allReturnsInGroup = new Dictionary<int, ReturnAccount>();
        Dictionary<int, ReturnAccount> settledReturnsInGroup = new Dictionary<int, ReturnAccount>();
        Dictionary<int, ReturnAccount> bankosReturnsInGroup = new Dictionary<int, ReturnAccount>();

        // SQL
        List<string> creditors = new List<string>();
        List<string> passFailClientRecordList = new List<string>();
        List<string> emailedToList = new List<string>();
        Boolean needsUpdate;
        int totalFilesEmailed = 0;

        Dictionary<int, int> monthsDays = new Dictionary<int, int>()
        {
            { 1, 31 },
            { 2, leap },
            { 3, 31 },
            { 4, 30 },
            { 5, 31 },
            { 6, 30 },
            { 7, 31 },
            { 8, 31 },
            { 9, 30 },
            { 10, 31 },
            { 11, 30 },
            { 12, 31 },
        };

        public Data()
        { // for critical process tracking
            try
            {
                bhConnStr = @"FILEDSN=G:\Instructions\Progress DSN\db1_uncommitted.dsn";

                List<string> frequenciesToRun = new List<string>();
                frequenciesToRun = NeedToRunToday();

                Console.WriteLine("Validating current date...");
                foreach (string frequencyType in frequenciesToRun) // checks logs to get list of types that need to run today (almost never > 1) and loops through list so the program is run for each frequency that needs to be run today (month and week)
                {
                    creditorGroupDict.Clear();  //empty out the dictionary, so the weeklies won't interfere with the monthlies

                    if (frequencyType != null && frequencyType != "") // if frequencyType string is not null or empty...
                    {
                        Console.WriteLine("Pulling all status codes from Bloodhound...");
                        SetBhStatuses(); // stores all possible status codes in table to be checked against later for validation

                        Console.WriteLine("Processing creditor groups in database...");
                        GetSQLData(); // Get all info from mySQL table for each creditor group and store in object, as well as list of objects
                        ClientGroupWorkFlow(frequencyType); // Loop through all creditors, calling neccessary methods on each; pass in current frequency

                        Console.WriteLine("Logging records...");
                        CreateLogRecord(creditorGroup, passFailClientRecordList, totalFilesEmailed, frequencyType); // stores records from the automated program in the database

                        CreateReturnSummary(frequenciesToRun); // generate a brief report in pdf format showing what the program processed and sent; email and save to universal folder
                    }
                }

                // write success file
                WriteLogFile(true);
            }
            catch (Exception ex)
            {
                CreateLogRecord(creditorGroup, passFailClientRecordList, totalFilesEmailed, "ERROR", ex.ToString()); // stores records from the automated program in the database

                // write failure file
                WriteLogFile(false, (string)ex.ToString());

                //throw ex;
            }
        }

        private void ClientGroupWorkFlow(string currentFrequency) // calls functions for each client in loop and is passed either monthly or weekly string
        {
            foreach (ClientGroup creditorGroup in creditorGroupDict.Values) // loops through ALL creditor groups and runs the below functions
            {
                Console.WriteLine("Searching Bloodhound for creditor numbers...");
                SetBhCreditorNumbers(creditorGroup); // sets creditor numbers for creditor group from BH for validation before moving forward

                Console.WriteLine("Checking creditor return frequencies and comparing run dates...");
                if (CompareDates(creditorGroup, currentFrequency)) // uses frequency and current date to check if the group should be ran/given a report
                { // CompareDates also contains a safety check in case the DB is not being updated

                    Console.WriteLine("Pulling in account data...");
                    GetAcctInfo(creditorGroup); // gets all returns data from BH for the specified date (adds each group's data to the BhDt datatable)

                    Console.WriteLine("Updating data...");
                    if (bhDt.Rows.Count > 0) // if data was pulled from Bloodhound...
                    {
                        SeparateReturns(creditorGroup); // returns a list of formatted data tables to loop through in Excel from one list of accounts
                        ProcessData(creditorGroup); // creates excel files and sheets, emails them and logs in SQL record table
                    }
                    UpdateRunDate(creditorGroup); // updates the last date run in mySQL DB if the group is run/processed/made an excel sheet for
                    passFailClientRecordList.Add(ClientGroup.GetCredGroup(creditorGroup)); // Adds creditor group to processed list for log table in SQL

                    Console.WriteLine("Cleaning up after ourselves...");
                    allReturnsInGroup.Clear();
                    bankosReturnsInGroup.Clear();
                    settledReturnsInGroup.Clear();

                    Console.WriteLine("On to the next creditor group...");
                }
            }
        }

        private List<string> NeedToRunToday() // checks if the program has already run today for weekly/monthly to prevent duplicates from running
        {
            int runAsDateCol = 2; // column position in SQL data table
            int frequencyCol = 3; // column position in SQL data table

            // list that will be returned containing any return types that still need to be run through
            List<string> runTheseFrequencies = new List<string>();
            // run types that need to be looped through
            runTheseFrequencies = ParseLogs();

            // if it is both the first of the month AND a Saturday, the list should already contain and keep MONTHLY and WEEKLY if they haven't already run today
            if (curDt.Day == 1 && curDt.DayOfWeek == DayOfWeek.Saturday)
            {
                // don't remove anything!
            }

            // if it is any day that isn't the first of the month but is saturday, we need to remove the MONTHLY frequency from the list if it is on the list
            else if (curDt.Day != 1 && curDt.DayOfWeek == DayOfWeek.Saturday)
            {
                while (runTheseFrequencies.Contains("MONTHLY"))
                {
                    runTheseFrequencies.Remove("MONTHLY");
                }
            }

            // if it is any day that isn't a Saturday but is the first, we need to remove the WEEKLY frequency from the list if it is on the list
            else if (curDt.DayOfWeek != DayOfWeek.Saturday && curDt.Day == 1)
            {
                while (runTheseFrequencies.Contains("WEEKLY"))
                {
                    runTheseFrequencies.Remove("WEEKLY");
                }
            }

            // if it's neither the first of the month or a saturday then we should take out both frequencies from the list
            else
            {
                while (runTheseFrequencies.Contains("WEEKLY"))
                {
                    runTheseFrequencies.Remove("WEEKLY");
                }
                while (runTheseFrequencies.Contains("MONTHLY"))
                {
                    runTheseFrequencies.Remove("MONTHLY");
                }
            }

            // remove any junk in the list (e.g. errors)
            foreach (string value in runTheseFrequencies)
            {
                if (value != "MONTHLY" && value != "WEEKLY")
                {
                    runTheseFrequencies.Remove(value);
                }
            }

            return runTheseFrequencies.Distinct().ToList(); // returns list of all return types that were NOT found in the logs for today that need to be run!
        }

        private void SetBhStatuses() // stores all possible statuses with their descriptions from BH in a data table for data validation later
        {
            creditorGroupDict.Clear();
            string checkStatusQuery = @"SELECT DISTINCT smstatcode, smstatdesc FROM PUB.statmstr"; // gets every status and description from BH

            using (OdbcConnection cnBH = new OdbcConnection(bhConnStr))
            {
                cnBH.Open();
                using (OdbcCommand command = new OdbcCommand(checkStatusQuery, cnBH))
                {
                    OdbcDataAdapter da = new OdbcDataAdapter(command);

                    if (da.ToString().Length > 0)
                    {
                        da.Fill(validStatuses); // fills data table with all statuses and descriptions
                    }
                }

                for (int status = 0; status < validStatuses.Rows.Count; status++) // sets statuses for bankos by description and status start letter
                {
                    if ((validStatuses.Rows[status][1].ToString().ToUpper().Contains("BANKRUPT") || validStatuses.Rows[status][1].ToString().ToUpper().Contains("BNK") || validStatuses.Rows[status][1].ToString().ToUpper().Contains("BANK") || validStatuses.Rows[status][1].ToString().ToUpper().Contains("BANKO") || validStatuses.Rows[status][1].ToString().ToUpper().Contains("BANKRUPTCY")) && validStatuses.Rows[status][0].ToString().ToUpper().StartsWith("R"))
                    {
                        bankoStatusCodes.Add(validStatuses.Rows[status][0].ToString().ToUpper()); // stores statuses that meet requirements in banko list
                    }
                }
            }

            // removes any blank status's row from the data table
            for (int y = 0; y < validStatuses.Rows.Count; y++)
            {
                if (validStatuses.Rows[y][0].ToString() == "")
                {
                    validStatuses.Rows.RemoveAt(y);
                    y--;
                }
            }

            // if no statuses were found, something is wrong...
            if (validStatuses.Rows.Count <= 0)
            {
                throw new Exception("No statuses were found in BloodHound!");
            }
        }

        private void GetSQLData() // Get all client data from DB, stores in data table, then creates objects for each row that are stored in a dictionary
        {
            System.Data.DataTable table = new System.Data.DataTable();
            using (OdbcConnection cnSQL = new OdbcConnection(sqlConnStr))
            {
                cnSQL.Open();
                selectSQL = "SELECT * FROM `Returns` WHERE `ID` >= 0 "; // Gets all values from data table where ID is at least zero (all)

                using (OdbcCommand selectCMD = new OdbcCommand(selectSQL, cnSQL))
                {
                    OdbcDataAdapter da = new OdbcDataAdapter(selectCMD);

                    if (da.ToString().Length > 0)
                    {
                        da.Fill(table);  // adds all creditor group data to sqDt DataTable
                    }
                }
            }

            // loops through all rows in the creditor info data table
            for (int b = 0; b < table.Rows.Count; b++)
            {
                // stores null values as empty strings to prevent errors later
                for (int a = 0; a < table.Columns.Count; a++)
                {
                    if (table.Rows[b][a] == System.DBNull.Value)
                    {
                        table.Rows[b][a] = "";
                    }
                }

                // removes duplicate creditors pulled from database
                Hashtable hTable = new Hashtable();
                if (hTable.Contains(table.Rows[b]))
                {
                    table.Rows.RemoveAt(b);
                }
                else
                {
                    hTable.Add(table.Rows[b], string.Empty);
                }
            }

            // Start fresh loop through all creditor rows and store client groups as objects
            for (int y = 0; y < table.Rows.Count; y++)
            {

                DataRow row = table.Rows[y]; // Sets row variable to equal the row in sqDt at the looped value

                ClientGroup creditorGroup = new ClientGroup( // create object for current client group
                    (int)row[0],
                    (string)row[1],
                    (string)row[2],
                    (string)row[3],
                    (string)row[4],
                    (string)row[5],
                    (string)row[6],
                    (string)row[7],
                    (string)row[8],
                    (string)row[9],
                    (string)row[10],
                    (string)row[11],
                    (string)row[12],
                    (string)row[13],
                    (string)row[14],
                    (string)row[15],
                    (string)row[16],
                    (string)row[17],
                    (string)row[18],
                    (string)row[19],
                    (string)row[20],
                    (string)row[21],
                    (string)row[22]);

                creditorGroupDict.Add(((string)row[2]), creditorGroup); // Add each object to the creditor group dictionary by their group
            }
        }

        private void SetBhCreditorNumbers(ClientGroup creditorGroup) // stores creditor numbers for creditor group from BH in data table for validation
        {
            bhCreditors.Clear(); // needs to clear data table for creditors because this function is called in a loop for all the creditors

            using (OdbcConnection cnBH = new OdbcConnection(bhConnStr))
            {
                cnBH.Open();
                string checkCredNumQuery = @"SELECT gdcnumber FROM PUB.credgrpd WHERE gdgnumber = ?"; // gets every creditor number for the creditor group

                using (OdbcCommand command = new OdbcCommand(checkCredNumQuery, cnBH))
                {
                    command.Parameters.Add("@grp", OdbcType.VarChar).Value = (string)ClientGroup.GetCredGroup(creditorGroup);
                    OdbcDataAdapter da = new OdbcDataAdapter(command);

                    if (da.ToString().Length > 0)
                    {
                        da.Fill(bhCreditors); // stores all creditor numbers for this one creditor group in data table
                    }
                }
            }
            if (bhCreditors.Rows.Count <= 0) // if no creditor numbers were found for this creditor group...
            {
                throw new Exception("Creditor group " + (string)ClientGroup.GetCredGroup(creditorGroup) + " has no creditor numbers in BloodHound!");
            }
        }

        private Boolean CompareDates(ClientGroup creditorGroup, string currentFreq) // takes ONE creditor group at a time and checks if it needs to be updated
        {
            needsUpdate = false; // set back to default - variable determines whether or not current client should have returns created today

            if (validator) // if validation is turned on then check if database has not been updated in a long time (might not be working)
            {
                if (Convert.ToDateTime(ClientGroup.GetCredLastRunDate(creditorGroup)) < curDt.AddDays(-70)) // if last run date is more than 70 days ago (in regards to adjusted curDt) then throw an exception to stop the program since it's likely not updating the DB correctly anyways
                {
                    // note: usually monthly clients will have a last run date of ~60 days ago, so we do not want to check for less than that
                    throw new Exception("Last run date is more than 70 days ago! SQL database may not be updating correctly. Please rerun backdated to avoid skipping or overlapping return records.");
                }
            }

            int curMonth = curDt.Month; // the current month (number value)
            int curMonthDays = GetDaysKey(curMonth); // the number of days in the current month
            
            switch ((string)ClientGroup.GetCredFrequency(creditorGroup)) // compare parameter frequency with client group frequency to set correct dates
            {
                case "MONTHLY":
                    if (currentFreq == "MONTHLY") // if the frequency we are currently running matches the client group's specifications...
                    {
                        firstDate = curDt.AddDays(-curMonthDays); // sets first date to the first day of the of the month
                        secondDate = curDt.AddDays(-1); // sets second date to yesterday which would be the last day of the month
                        string lastRun = ClientGroup.GetCredLastRunDate(creditorGroup);
                        if (Convert.ToDateTime(lastRun).Date <= firstDate.Date)
                        {
                            needsUpdate = true; // the client needs to have new return files generated so we return true to continue in the program
                        }
                    }
                    break;
                case "WEEKLY":
                    if (currentFreq == "WEEKLY") // if the frequency we are currently running matches the client group's specifications...
                    {
                        if (validator) // if validation is turned on then check if database has not been updated in a long time (might not be working)
                        {
                            if (Convert.ToDateTime(ClientGroup.GetCredLastRunDate(creditorGroup)) < curDt.AddDays(-15)) // if last run date is more than 15 days ago (in regards to adjusted curDt) then throw an exception to stop the program since it's likely not updating the DB
                            {
                                throw new Exception("Last run date is more than 15 days ago for this weekly-cycle client! SQL database may not be updating correctly. Please rerun backdated to avoid skipping or overlapping return records.");
                            }
                        }

                        firstDate = curDt.AddDays(-7).Date; // sets first date to a week ago (should be a friday)
                        secondDate = curDt.AddDays(-1).Date; // sets second date to yesterday (should be a friday)
                        if (Convert.ToDateTime(ClientGroup.GetCredLastRunDate(creditorGroup)) <= firstDate.Date)
                        {
                            needsUpdate = true; // the client needs to have new return files generated so we return true to continue in the program
                        }
                    }
                    break;
                default:
                    // if the creditor group's frequency value is neither MONTHLY or WEEKLY...
                    throw new Exception("Invalid creditor group frequency: " + (string)ClientGroup.GetCredFrequency(creditorGroup));
            }

            return needsUpdate;
        }

        private void GetAcctInfo(ClientGroup creditorGroup) // queries bloodhound for return accounts within the specified dates for current cred group
        {
            bhDt.Clear(); // clears out data table that contains return account from bloodhound by creditor group (is part of creditor loop)

            using (OdbcConnection cnBH = new OdbcConnection(bhConnStr))
            {
                cnBH.Open();
                if (((string)ClientGroup.GetCredGroup(creditorGroup) == "RGP-G") || ((string)ClientGroup.GetCredGroup(creditorGroup) == "MRG-G"))
                // if creditor group is RGP-G or MRG-G
                {
                    selectSQL = (@"SELECT amlinkacct AS 'Linked', amstatus AS 'Status', gdgnumber as 'Creditor', amdnumber as 'Debtor Number', amclacctno AS 'Your Account #', amakaname AS 'Guarantor Name', ampatient AS 'Patient Name', smstatdesc as 'Return Description', amretnamt as 'Return Amount', amamtlstd as 'Amount Listed', amretndate as 'Return Date', amlastdate as 'Last Pay Date', amlastamt as 'Last Pay Amt', amcnumber AS 'Creditor Number', amlstactdt as 'Date of Service' FROM PUB.acctmstr INNER JOIN PUB.statmstr ON PUB.acctmstr.amstatus = PUB.statmstr.smstatcode INNER JOIN PUB.credgrpd ON PUB.acctmstr.amcnumber = PUB.credgrpd.gdcnumber WHERE PUB.credgrpd.gdgnumber = ? AND amretndate >= ? AND amretndate <= ? AND amstatus <> 'LA2' ORDER BY amstatus");
                }
                else if ((string)ClientGroup.GetCredGroup(creditorGroup) == "VMG-G")
                // if creditor group is VMG-G (leaves out deceased, etc)
                {
                    selectSQL = (@"SELECT amlinkacct AS 'Linked', amstatus AS 'Status', gdgnumber as 'Creditor', amdnumber as 'Debtor Number', amclacctno AS 'Your Account #', amakaname AS 'Guarantor Name', ampatient AS 'Patient Name', smstatdesc as 'Return Description', amretnamt as 'Return Amount', amamtlstd as 'Amount Listed', amretndate as 'Return Date', amlastdate as 'Last Pay Date', amlastamt as 'Last Pay Amt', amcnumber AS 'Creditor Number', amlstactdt as 'Date of Service' FROM PUB.acctmstr INNER JOIN PUB.statmstr ON PUB.acctmstr.amstatus = PUB.statmstr.smstatcode INNER JOIN PUB.credgrpd ON PUB.acctmstr.amcnumber = PUB.credgrpd.gdcnumber WHERE PUB.credgrpd.gdgnumber = ? AND amretndate >= ? AND amretndate <= ? AND amstatus <> 'LA2' AND amstatus <> '3V' AND amstatus <> 'S' ORDER BY amstatus");
                }
                else
                // default value for all other creditor groups
                {
                    selectSQL = (@"SELECT amlinkacct AS 'Linked', amstatus AS 'Status', gdgnumber as 'Creditor', amdnumber as 'Debtor Number', amclacctno AS 'Your Account #', amakaname AS 'Guarantor Name', ampatient AS 'Patient Name', smstatdesc as 'Return Description', amretnamt as 'Return Amount', amamtlstd as 'Amount Listed', amretndate as 'Return Date', amlastdate as 'Last Pay Date', amlastamt as 'Last Pay Amt', amcnumber AS 'Creditor Number', amlstactdt as 'Date of Service' FROM PUB.acctmstr INNER JOIN PUB.statmstr ON PUB.acctmstr.amstatus = PUB.statmstr.smstatcode INNER JOIN PUB.credgrpd ON PUB.acctmstr.amcnumber = PUB.credgrpd.gdcnumber WHERE PUB.credgrpd.gdgnumber = ? AND amretndate >= ? AND amretndate <= ? AND amstatus <> 'LA2' ORDER BY amstatus");
                }

                using (OdbcCommand command = new OdbcCommand(selectSQL, cnBH))
                {
                    command.Parameters.Add("@grp", OdbcType.VarChar).Value = (string)ClientGroup.GetCredGroup(creditorGroup); // greditor group parameter
                    command.Parameters.Add("@firstMonthDate", OdbcType.Date).Value = firstDate.ToShortDateString(); // start date range parameter
                    command.Parameters.Add("@secondMonthDate", OdbcType.Date).Value = secondDate.ToShortDateString(); // end date range parameter
                    OdbcDataAdapter da = new OdbcDataAdapter(command);

                    if (da.ToString().Length > 0) // if accounts/results were found...
                    {
                        da.Fill(bhDt);  // adds all return account records from BH to DataTable

                        for (int a = 0; a < bhDt.Columns.Count; a++) // loops through columns
                        {
                            for (int b = 0; b < bhDt.Rows.Count; b++) // loops through rows
                            {
                                if (bhDt.Rows[b][a] == System.DBNull.Value || bhDt.Rows[b][a] == null || bhDt.Rows[b][a].ToString() == "") // if cell null
                                {
                                    bhDt.Rows[b][a] = ""; // Sets null values in DT to empty strings to prevent errors later on
                                }
                            }
                        }
                    }
                    da.Dispose();
                }

                FilterOutSharedAccounts(); // removes all shared accounts from data table that don't have LA2 status

                System.Data.DataTable bhDtBanko = new System.Data.DataTable(); // temp table for banko information only

                String selectBankoBh = (@"select nullif(pro_arr_descape(pro_element(wdtext,9,16)),'?') from PUB.windata where wdnumber = ? AND wdtype = 'D' AND wdcode = 'P'"); // get all banko window information for the specified debtor number (parameter)

                using (OdbcCommand command = new OdbcCommand(selectBankoBh, cnBH))
                {
                    if (bhDt.Rows.Count > 0) // if there are return accounts that need to be included in the return file(s)...
                    {
                        for (int y = 0; y < bhDt.Rows.Count; y++) // loops through every row (every return account) in data table
                        {
                            command.Parameters.Add("@debtNumber", OdbcType.VarChar).Value = (string)bhDt.Rows[y][2]; // sends in debtor number parameter

                            OdbcDataReader reader = command.ExecuteReader();
                            while (reader.Read())
                            {
                                String temp = reader.GetString(0); // gets all banko values for debtor number as string
                                temp = temp + ";" + (string)bhDt.Rows[y][2]; // adds debtor number to string of values that will be split into array
                                string[] values = temp.Split(';'); // splits all banko values for debtor number as string

                                if (values[0].Length > 1) // checks if debtor actually has banko values stored by checking the first value (case number) which is the crucial value for bankos (if values in array are blank they will still have empty strings)...
                                {
                                    // insert a column in banko data table for every value in the array
                                    while (bhDtBanko.Columns.Count < values.Length)
                                    {
                                        DataColumn newCol = new DataColumn();
                                        bhDtBanko.Columns.Add(newCol);
                                    }

                                    // set name of id column in banko DT
                                    bhDtBanko.Columns[bhDtBanko.Columns.Count - 1].ColumnName = "Debtor Number";

                                    // insert data from array into banko data table by adding new rows for each debtor number
                                    if (values.Length > 0 && values[0] != null && values[0] != "")
                                    {
                                        DataRow newRow = bhDtBanko.NewRow();
                                        for (int j = 0; j < values.Length; j++)
                                        {
                                            newRow[j] = values[j].Trim();
                                        }
                                        bhDtBanko.Rows.Add(newRow);
                                    }
                                }
                            }
                            command.Parameters.Clear();
                            reader.Close(); // important to prevent error
                        }
                    }

                    // if banko information exists in the banko table...
                    if (bhDtBanko.Rows.Count > 0 && bhDtBanko.Rows[0][0].ToString().Length > 1)
                    {
                        // eliminate duplicate banko info rows
                        Hashtable bankoWindow = new Hashtable();

                        for (int row = 0; row < bhDtBanko.Rows.Count; row++)
                        {
                            if (!(bankoWindow.ContainsKey(bhDtBanko.Rows[row][bhDtBanko.Columns.Count - 1].ToString())))
                            {
                                bankoWindow.Add(bhDtBanko.Rows[row][bhDtBanko.Columns.Count - 1].ToString(), bhDtBanko.Rows[row][0].ToString());
                            }
                            else
                            {
                                bhDtBanko.Rows.RemoveAt(row);
                            }
                        }

                        // change any null values to empty strings to prevent errors when using objects/classes
                        for (int a = 0; a < bhDtBanko.Columns.Count; a++)
                        {
                            for (int b = 0; b < bhDtBanko.Rows.Count; b++)
                            {
                                if (bhDtBanko.Rows[b][a] == System.DBNull.Value || bhDtBanko.Rows[b][a] == null || bhDtBanko.Rows[b][a].ToString() == "" || !bhDtBanko.Rows[b][a].ToString().Contains("#"))
                                {
                                    bhDtBanko.Rows[b][a] = "";
                                }
                            }
                        }

                        // add banko columns in banko table to the bhDt table if they are NOT already in the table
                        if (MissingBankoData(bhDt))
                        {
                            // loop through banko columns to be added to the main data table if they're missing from it
                            for (int bankoCols = 0; bankoCols < bankoColNames.Length; bankoCols++)
                            {
                                DataColumn newCol = new DataColumn();
                                newCol.ColumnName = bankoColNames[bankoCols];
                                bhDt.Columns.Add(newCol);
                            }
                        }

                        // pass the two tables into the function that will combine them into one: bhDt
                        CombineDts(bhDt, bhDtBanko);

                        // clear for next client group/next time function called
                        bhDtBanko.Clear();
                    }
                }

                // create an object for each return account for this creditor group
                if (!MissingBankoData(bhDt)) // has banko columns
                {
                    for (int t = 0; t < bhDt.Rows.Count; t++) // loops through each row/account to store the row/account as an object
                    {
                        ReturnAccount currentAccountObj = new ReturnAccount(
                                                // create new object for current account's info including banko info
                                                Convert.ToString(bhDt.Rows[t][0]),
                                                Convert.ToString(bhDt.Rows[t][1]),
                                                Convert.ToString(bhDt.Rows[t][2]),
                                                Convert.ToString(bhDt.Rows[t][3]),
                                                Convert.ToString(bhDt.Rows[t][4]),
                                                Convert.ToString(bhDt.Rows[t][5]),
                                                Convert.ToString(bhDt.Rows[t][6]),
                                                Convert.ToString(bhDt.Rows[t][7]),
                                                Convert.ToString(bhDt.Rows[t][8]),
                                                Convert.ToString(bhDt.Rows[t][9]),
                                                Convert.ToString(bhDt.Rows[t][10]),
                                                Convert.ToString(bhDt.Rows[t][11]),
                                                Convert.ToString(bhDt.Rows[t][12]),
                                                Convert.ToString(bhDt.Rows[t][13]),
                                                Convert.ToString(bhDt.Rows[t][14]),
                                                Convert.ToString(bhDt.Rows[t][15]),
                                                Convert.ToString(bhDt.Rows[t][16]),
                                                Convert.ToString(bhDt.Rows[t][17]),
                                                Convert.ToString(bhDt.Rows[t][18]));

                        // validate creditor number and status, then add account object to main account dictionary
                        if (StatusValid(ReturnAccount.GetReturnStatus(currentAccountObj).ToString()) && CreditorNumberValid(ReturnAccount.GetReturnCredNumber(currentAccountObj).ToString()))
                        {
                            allReturnsInGroup.Add(t, currentAccountObj);
                        }
                        else // throw an exception if the status or creditor number cannot be validated
                        {
                            throw new Exception("Cannot validate return account status code: " + ReturnAccount.GetReturnStatus(currentAccountObj).ToString() + " or creditor number: " + ReturnAccount.GetReturnCredNumber(currentAccountObj).ToString());
                        }
                    }
                }
                else // does not contain banko info because the bhDt table does not contain a case number column
                {
                    for (int t = 0; t < bhDt.Rows.Count; t++) // loops through each row/account to create an object for each return account from table
                    {
                        ReturnAccount currentAccountObj = new ReturnAccount(
                                                // create new object for current account/row
                                                Convert.ToString(bhDt.Rows[t][0]),
                                                Convert.ToString(bhDt.Rows[t][1]),
                                                Convert.ToString(bhDt.Rows[t][2]),
                                                Convert.ToString(bhDt.Rows[t][3]),
                                                Convert.ToString(bhDt.Rows[t][4]),
                                                Convert.ToString(bhDt.Rows[t][5]),
                                                Convert.ToString(bhDt.Rows[t][6]),
                                                Convert.ToString(bhDt.Rows[t][7]),
                                                Convert.ToString(bhDt.Rows[t][8]),
                                                Convert.ToString(bhDt.Rows[t][9]),
                                                Convert.ToString(bhDt.Rows[t][10]),
                                                Convert.ToString(bhDt.Rows[t][11]),
                                                Convert.ToString(bhDt.Rows[t][12]),
                                                Convert.ToString(bhDt.Rows[t][13]));

                        // validate status and creditor number, then add the account object to the dictionary of all return accounts
                        if (StatusValid(ReturnAccount.GetReturnStatus(currentAccountObj).ToString()) && CreditorNumberValid(ReturnAccount.GetReturnCredNumber(currentAccountObj).ToString()))
                        {
                            allReturnsInGroup.Add(t, currentAccountObj);
                        }
                        else // if status or creditor number cannot be validated, throw an error
                        {
                            throw new Exception("Cannot validate return account status code: " + ReturnAccount.GetReturnStatus(currentAccountObj).ToString() + " or creditor number: " + ReturnAccount.GetReturnCredNumber(currentAccountObj).ToString());
                        }
                    }
                }
            }
        }

        private List<System.Data.DataTable> SeparateReturns(ClientGroup creditorGroup) // returns a list of formatted data tables to loop through in Excel
        {
            bool distinctReturns = true; // determines whether client wants to include or exclude split returns in their returns file or sheet(s)

            if (ClientGroup.GetCredDistinctReturns(creditorGroup) == "NO") // value No means distinctReturns should be False
            {
                distinctReturns = false; // client wants accounts separated into other files or sheets to also be in the returns file or sheet
            }

            splitReturnTablesList.Clear(); // clear out main list (that will be returned from function) of data tables from last client group iteration

            // itemized clients
            if (ClientGroup.GetCredItemized(creditorGroup) == "YES") // client would like separate files for bankos, settleds, and returns
            {
                // custom splitting
                // if (ClientGroup.GetCredCustomSplit(creditorGroup) == "YES" && ClientGroup.GetCredSplitType(creditorGroup) == "ITEMIZED") 

                // standard itemized splitting
                foreach (ReturnAccount rtrn in allReturnsInGroup.Values.ToList()) // loops through every acct ID for BH query results
                {
                    // Standard file splitting - should have three files (banko, settled, returns) + custom splitting
                    if ((ReturnAccount.GetReturnStatus(rtrn) == "7") || ((ReturnAccount.GetReturnStatus(rtrn) == "Z")) && (ReturnAccount.GetReturnDescription(rtrn).Contains("SETTLED")) || (ReturnAccount.GetReturnDescription(rtrn).Contains("SETTLED IN FULL"))) // account is a Settled In Full
                    {
                        SwitchReturnLists(rtrn, allReturnsInGroup, settledReturnsInGroup, distinctReturns); // moves or copies from all returns list to settleds list
                    }
                    else if ((ReturnAccount.GetReturnDescription(rtrn).Contains("CANCEL BNK") || ReturnAccount.GetReturnDescription(rtrn).Contains("CANCEL BANK") || ReturnAccount.GetReturnDescription(rtrn).Contains("BANKRUPTCY")) || ReturnAccount.GetReturnDescription(rtrn).Contains("CANCEL BANKRUPT") || bankoStatusCodes.Contains(ReturnAccount.GetReturnStatus(rtrn).ToString()) && ReturnAccount.GetReturnBhCode(rtrn) != "") // description contains bankruptcy in some form, and there is a case number/banko code in bloodhound
                    {
                        SwitchReturnLists(rtrn, allReturnsInGroup, bankosReturnsInGroup, distinctReturns); // moves or copies from all returns list to bankos list
                    }
                }
                // Create data table for each file/split and add to final return list for excel
                if (bankosReturnsInGroup.Count > 0) { splitReturnTablesList.Add(ReturnsListToTable(bankosReturnsInGroup, "Bankruptcy", creditorGroup)); };
                if (settledReturnsInGroup.Count > 0) { splitReturnTablesList.Add(ReturnsListToTable(settledReturnsInGroup, "Settled in Full", creditorGroup)); };
                if (allReturnsInGroup.Count > 0) { splitReturnTablesList.Add(ReturnsListToTable(allReturnsInGroup, "Returns", creditorGroup)); };
            }

            // breakout clients
            else if (ClientGroup.GetCredBreakout(creditorGroup) == "YES") // client would like separate Excel sheets for separate creditor numbers
            {
                // custom splitting
                if (ClientGroup.GetCredCustomSplit(creditorGroup) == "YES" && ClientGroup.GetCredSplitType(creditorGroup) == "BREAKOUT")
                {
                    List<Dictionary<int, ReturnAccount>> acctDictionariesList = new List<Dictionary<int, ReturnAccount>>(); // for custom-split dictionaries
                    acctDictionariesList = (CustomReturnSplitting(creditorGroup, distinctReturns)).ToList(); // calls function to do custom splitting for all accounts

                    foreach (Dictionary<int, ReturnAccount> acctDictionary in acctDictionariesList.ToList()) // loops through all dictionaries in list of custom splits
                    {
                        if (acctDictionary.Keys.Count > 0) // if dictionary is not empty...
                        {
                            string currentCredNumber = ""; // default value
                            string currentStatusCode = ""; // default value

                            foreach (ReturnAccount acctInDict in acctDictionary.Values.ToList()) // loops through accounts in dictionary once to get cred number and status
                            {
                                currentCredNumber = (string)ReturnAccount.GetReturnCredNumber(acctInDict);
                                currentStatusCode = (string)ReturnAccount.GetReturnStatus(acctInDict);
                                break; // just need one cred number since they're all the same in the table; need values for next if statement
                            }
                            if (acctDictionary.Count > 0) // if account dictionary for custom splits is not empty...
                            {
                                splitReturnTablesList.Add(ReturnsListToTable(acctDictionary, (currentCredNumber + "_" + currentStatusCode), creditorGroup)); // creates a datatable for each custom split type dictionary of return account objects and adds to final list of returned tables for excel
                            }
                        }
                    }
                }

                // standard breakout splitting - should have one table for each creditor number + custom splits
                Dictionary<int, ReturnAccount> newDict = new Dictionary<int, ReturnAccount>();

                for (int cell = 0; cell < bhCreditors.Rows.Count; cell++) // loops through creditor numbers for group
                {
                    foreach (int keyVal in allReturnsInGroup.Keys.ToList()) // loops through every acct ID for BH query results per creditor number
                    {
                        if ((string)ReturnAccount.GetReturnCredNumber(allReturnsInGroup[keyVal]) == bhCreditors.Rows[cell][0].ToString()) // if creditor number on account matches current creditor number in loop
                        {
                            SwitchReturnLists(allReturnsInGroup[keyVal], allReturnsInGroup, newDict, distinctReturns); // populate new dictionary with account
                        }
                    }

                    if (newDict.Count > 0) { splitReturnTablesList.Add(ReturnsListToTable(newDict, bhCreditors.Rows[cell][0].ToString(), creditorGroup)); }; // converts list of accounts for cred number to data table and adds table to main list

                    newDict.Clear(); // clears creditor number dictionary for next cred number in loop
                }
            }
            return splitReturnTablesList; // returns list of all tables stored
        }

        private void ProcessData(ClientGroup clientGroup) // for current client group, creates correct type of Excel file(s) and then emails file to specified address
        {
            // if the returns list with tables for Excel is not empty...
            if (splitReturnTablesList.Count > 0)
            {
                // if client wants anything beyond only one file and only one sheet in the file with all returns then...
                if (ClientGroup.GetCredBreakout(clientGroup) == "YES" || ClientGroup.GetCredItemized(clientGroup) == "YES" || ClientGroup.GetCredCustomSplit(clientGroup) == "YES")
                {
                    // if client specifically wants a creditor number breakout/separate sheets...
                    if (ClientGroup.GetCredBreakout(clientGroup) == "YES")
                    {
                        // create Excel file and email
                        Email(splitReturnTablesList[0].TableName, BreakoutExcelOutput(splitReturnTablesList.ToList(), clientGroup), clientGroup);
                    }

                    // if client wants only one Excel sheet within a file...
                    else if (ClientGroup.GetCredBreakout(clientGroup) == "NO")
                    {
                        // loop through all tables in the returns list...
                        foreach (System.Data.DataTable returnTable in splitReturnTablesList.ToList())
                        {
                            // emails to specified address with Excel file created attached
                            Email(returnTable.TableName, ItemizedExcelOutput(returnTable, returnTable.TableName, clientGroup), clientGroup);
                        }
                    }
                }
            }

            // if client wants only one file and only one sheet in the file with all returns then...
            else if (ClientGroup.GetCredBreakout(clientGroup) == "NO" && ClientGroup.GetCredItemized(clientGroup) == "NO" && ClientGroup.GetCredCustomSplit(clientGroup) == "NO")
            {
                // create the Excel file and email it out
                Email("Returns", (ItemizedExcelOutput(ReturnsListToTable(allReturnsInGroup, "Returns", clientGroup), "Returns", clientGroup)), clientGroup);
            }

            // if client doesn't want a single file with a single sheet OR (multiple files and/or multiple sheets)
            else
            {
                throw new Exception("Logic error or invalid data for breakout and itemization values. Check UpdateIfNeeded method.");
            }
        }

        private void UpdateRunDate(ClientGroup credGroup) // updates the last run date value in client group SQL table for passed-in client group
        {
            using (OdbcConnection cnSQL = new OdbcConnection(sqlConnStr))
            {
                cnSQL.Open();
                selectSQL = (@"UPDATE `Returns` SET `Last Run Date` = ? WHERE `Client Group` = ?"); // query for updating mySQL Last Run Date to current date

                using (OdbcCommand selectCMD = new OdbcCommand(selectSQL, cnSQL))
                {
                    selectCMD.Parameters.Add("@newDate", OdbcType.VarChar).Value = curDt.ToString("MM/dd/yyyy"); // formatted date
                    selectCMD.Parameters.Add("@creditorGroup", OdbcType.VarChar).Value = ClientGroup.GetCredGroup(credGroup); // paramater creditor group
                    selectCMD.ExecuteNonQuery();
                }
            }
        }

        private void CreateLogRecord(ClientGroup clientGroup, List<string> passFailClientRecords, int totalFiles, string frequency, string error = "PASSED") // creates new row in run logs table in SQL with record of running
        {
            // stores all clients processed in a list to instert into the DB
            StringBuilder clientSb = new StringBuilder();
            if (passFailClientRecordList.Count > 0)
            {
                for (int clientPos = 0; clientPos < passFailClientRecords.Count - 1; clientPos++)
                {
                    clientSb.Append(passFailClientRecords[clientPos] + ",");
                }
                clientSb.Append(passFailClientRecords[passFailClientRecords.Count - 1]); // add last client without a comma
            }
            else
            {
                clientSb.Append("");
            }

            // stores all email addresses sent to in a list to instert into the DB
            StringBuilder emailSb = new StringBuilder();
            if (emailedToList.Distinct().Count() > 0)
            {
                for (int email = 0; email < emailedToList.Distinct().Count() - 1; email++)
                {
                    emailSb.Append(emailedToList[email] + ",");
                }
                emailSb.Append(emailedToList[emailedToList.Count - 1]); // add last client without a comma
            }
            else
            {
                emailSb.Append("");
            }

            // store in DB
            using (OdbcConnection cnSQL = new OdbcConnection(sqlConnStr))
            {
                cnSQL.Open();
                selectSQL = (@"INSERT INTO `RunLogs` (`runDate`, `runAsDate`, `runType`, `clientsIncluded`, `emailed`, `totalFilesEmailed`, `errorMessage`) VALUES (?, ?, ?, ?, ?, ?, ?);");
                // query for updating mySQL Last Run Date to current date

                using (OdbcCommand selectCMD = new OdbcCommand(selectSQL, cnSQL))
                {
                    selectCMD.Parameters.Add("@newDate", OdbcType.Date).Value = ((DateTime)DateTime.Now).Date;
                    selectCMD.Parameters.Add("@newDate", OdbcType.Date).Value = ((DateTime)curDt).Date;
                    selectCMD.Parameters.Add("@frequency", OdbcType.VarChar).Value = (string)frequency; // global
                    selectCMD.Parameters.Add("@clientSb", OdbcType.VarChar).Value = clientSb.ToString(); // passed in
                    selectCMD.Parameters.Add("@emailSb", OdbcType.VarChar).Value = emailSb.ToString(); // global
                    selectCMD.Parameters.Add("@totalFiles", OdbcType.Int).Value = (int)totalFiles; // passed in
                    selectCMD.Parameters.Add("@error", OdbcType.VarChar).Value = (string)error; // passed in

                    selectCMD.ExecuteNonQuery();
                }
            }
        }

        #region Internal Functions
        private List<string> ParseLogs()
        {
            // gets log records from SQL table and stores them in a temporary data table
            System.Data.DataTable tempTable = GetLogDailyRecords();

            // create list for which types should be updated based on which ones were not updated today yet (does NOT check current date here)
            List<string> frequenciesRunToday = new List<string>();

            // create list for which types should be updated based on which ones were not updated today yet (does NOT check current date here)
            List<string> updateTheseTypes = new List<string>();

            // which column in the SQL table contains the run as date value
            int runAsDateCol = 2;

            // which column in the SQL table contains the frequency type value
            int frequencyCol = 3;

            // check if the program is running too many times for the same frequency or run date
            int monthlyDupes = 0;
            int weeklyDupes = 0;

            // if there was at least one log that has a rundate that equals the current run date...
            if (tempTable.Rows.Count > 0)
            {
                // loops through temporary table rows so we can set the runTypeRecord to reflect the frequency value in the SQL log table for comparison later
                for (int row = 0; row < tempTable.Rows.Count; row++)
                {
                    // add the frequency value to the list
                    frequenciesRunToday.Add(tempTable.Rows[row][frequencyCol].ToString().ToUpper());
                } 
            }

            // use the frequencies logged today (list) to determine which frequencies still need to be run
            if (!frequenciesRunToday.Contains("WEEKLY") && !frequenciesRunToday.Contains("MONTHLY"))
            {
                // if the frequency values listed today are neither monthly or weekly (e.g. errors), then add both to the run list...
                updateTheseTypes.Add("WEEKLY");
                updateTheseTypes.Add("MONTHLY");
            }

            // if the frequency values for today contains WEEKLY but not MONTHLY, we need to run MONTHLY...
            else if (frequenciesRunToday.Contains("WEEKLY") && !frequenciesRunToday.Contains("MONTHLY"))
            {
                updateTheseTypes.Add("MONTHLY");
            }

            // if the frequency values for today contains MONTHLY but not WEEKLY, we need to run WEEKLY...
            else if (frequenciesRunToday.Contains("MONTHLY") && !frequenciesRunToday.Contains("WEEKLY"))
            {
                updateTheseTypes.Add("WEEKLY");
            }

            // loop through list values to check for critical duplicates...
            foreach (string value in updateTheseTypes)
            {
                switch (value.Trim().ToUpper())
                {
                    case "MONTHLY":
                        monthlyDupes++;
                        break;
                    case "WEEKLY":
                        weeklyDupes++;
                        break;
                    default:
                        // remove any other values from list
                        updateTheseTypes.Remove(value);
                        break;
                }
            }

            // if there is more than one record for the current run date per frequency type...
            if (monthlyDupes > 1 || weeklyDupes > 1)
            {
                throw new Exception("This program was run more than once for monthly clients on the current run date! This could cause duplicate data.");
            }

            return updateTheseTypes;
        }

        private void CombineDts(System.Data.DataTable originalDt, System.Data.DataTable absorbedDt) // combines main table with all account data and banko-only table
        {
            for (int absRow = 0; absRow < absorbedDt.Rows.Count; absRow++) // loops through banko-only table rows
            {
                for (int i = 0; i < originalDt.Rows.Count; i++) // loops through main table rows regardless of the following:
                {
                    if ((absorbedDt.Rows[absRow][absorbedDt.Columns.Count - 1]).ToString() == (originalDt.Rows[i][2]).ToString()) // if debtor numbers in both tables match...
                    {
                        // add data from banko table to bhDt table in the same row where the matching debtor number is under the indexes of the banko-specific columns
                        originalDt.Rows[i][originalDt.Columns.Count - 5] = absorbedDt.Rows[absRow][0];
                        originalDt.Rows[i][originalDt.Columns.Count - 4] = absorbedDt.Rows[absRow][1];
                        originalDt.Rows[i][originalDt.Columns.Count - 3] = absorbedDt.Rows[absRow][2];
                        originalDt.Rows[i][originalDt.Columns.Count - 2] = absorbedDt.Rows[absRow][3];
                        originalDt.Rows[i][originalDt.Columns.Count - 1] = absorbedDt.Rows[absRow][4];
                    }
                }
            }
        }

        private bool StatusValid(string status) // validates status code through bloodhound
        {
            bool isValid = false; // default value

            for (int i = 0; i < validStatuses.Rows.Count; i++) // loops through all statuses in the valid statuses data table
            {
                if (validStatuses.Rows[i][0].ToString().ToUpper() == status.Trim().ToUpper()) // checks if the parameter status value matches any in the table
                {
                    isValid = true; // if there is a match the program will return true
                    break;
                }
            }
            return isValid; // otherwise the program will return false
        }

        private bool CreditorNumberValid(string number) // validates creditor number through bloodhound
        {
            bool isValid = false; // default value
            for (int i = 0; i < bhCreditors.Rows.Count; i++) // loops through all creditor numbers in the valid creditor number data table for the current client group
            {
                if (bhCreditors.Rows[i][0].ToString().ToUpper() == number.Trim().ToUpper()) // checks if the parameter creditor number value matches any in the table
                {
                    isValid = true; // if there is a match the program will return true
                    break;
                }
            }
            return isValid; // otherwise the program will return false
        }

        private int GetDaysKey(int month) // gets number of days in the specified month
        {
            int totalDays = 0; // default value

            foreach (KeyValuePair<int, int> curPair in monthsDays) // loops through every month
            {
                if (curPair.Key == month) // if the integer key matches the integer month number
                {
                    totalDays = curPair.Value; // store the value of the keypair as the totaldays
                    break; // stop running because it can only be one month at a time
                }
            }
            return totalDays; // return how many days are in the specified month
        }

        private string ItemizedExcelOutput(System.Data.DataTable table, string rType, ClientGroup credGroup) // for client groups that do not want multiple sheets in their excel files (one or more files) - called once per file
        {
            if (table != null && table.Rows.Count > 0 && credGroup != null && ClientGroup.GetCredBreakout(credGroup) == "NO") // if client wants only one sheet...
            {
                string fileName = (string)(SetExcelFileName(table, credGroup)); // calls function that sets the Excel file name from the table name
                filePath = ((string)ClientGroup.GetCredPath(credGroup) + fileName + ".xls"); // set full file path from filename and provided file directory path
                int count = 1; // duplicate file counter

                while ((File.Exists(filePath))) // while the file path (with filename) exists already...
                {
                    fileName = (fileName + "_" + count); // change the name to include a duplicate count at the end of the file name
                    filePath = ((string)ClientGroup.GetCredPath(credGroup) + fileName + ".xls"); // update the file path with this new name
                    count++; // add one to the dupe counter
                }

                // create excel file
                Console.WriteLine("Creating Excel file...");
                Microsoft.Office.Interop.Excel.Application xlApp;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                xlApp = new Microsoft.Office.Interop.Excel.Application();

                if (xlApp == null)
                {
                    throw new Exception("Excel is not properly installed!");
                }
                else
                {
                    xlApp.Visible = false;
                    xlApp.DisplayAlerts = false;
                    xlApp.ScreenUpdating = false;
                    object misValue = System.Reflection.Missing.Value;
                    var xlWorkBooks = xlApp.Workbooks;
                    xlWorkBook = xlWorkBooks.Add(misValue);
                    xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                    // column value formatting
                    xlWorkSheet.Columns["A:A"].NumberFormat = "@";
                    xlWorkSheet.Columns["J:J"].NumberFormat = "@";
                    xlWorkSheet.Columns["G:G"].NumberFormat = "mm/dd/yyyy";
                    xlWorkSheet.Columns["H:H"].NumberFormat = "mm/dd/yyyy";
                    xlWorkSheet.Columns["K:K"].NumberFormat = "mm/dd/yyyy";
                    xlWorkSheet.Columns["M:M"].NumberFormat = "mm/dd/yyyy";
                    xlWorkSheet.Columns["N:N"].NumberFormat = "mm/dd/yyyy";
                    xlWorkSheet.Columns["O:O"].NumberFormat = "mm/dd/yyyy";
                    xlWorkSheet.Columns["P:P"].NumberFormat = "mm/dd/yyyy";

                    // set row and col counts
                    int rows = table.Rows.Count;
                    int columns = table.Columns.Count;

                    // Add the +1 to allow room for column headers
                    var data = new object[rows + 1, columns];

                    // populate headers in excel sheet
                    int colHeaderCt = 0; // default value for number of headers in the sheet

                    foreach (DataColumn col in table.Columns) // loop through every column in the passed-in table
                    {
                        data[0, colHeaderCt] = (string)col.ColumnName; // set the specified cell (column from loop, first row only) to the table's column name
                        colHeaderCt++;
                    }

                    // populate data object for inserting into excel from datatable
                    for (int row = 0; row < rows; row++)
                    {
                        for (int col = 0; col < columns; col++)
                        {
                            if ((string)table.Rows[row][col] == "01/01/0001" || (string)table.Rows[row][col] == "01/01/1000" || (string)table.Rows[row][col] == "1/1/1000" || (string)table.Rows[row][col] == "1/1/0001") // if any of the cells are dates that are any of these placeholders...
                            {
                                data[row + 1, col] = ""; // change the Excel date value to an empty string
                            }
                            else // if they're not any of these dates...
                            {
                                data[row + 1, col] = table.Rows[row][col]; // add the value to the Excel object
                            }
                        }
                    }

                    // write this data to the excel worksheet
                    Range beginWrite = (Range)xlWorkSheet.Cells[1, 1];
                    Range endWrite = (Range)xlWorkSheet.Cells[rows + 1, columns];
                    Range sheetData = xlWorkSheet.Range[beginWrite, endWrite];
                    sheetData.Value2 = data;
                    sheetData.Select();

                    // page settings
                    xlApp.Cells.Font.Size = 11;
                    xlWorkSheet.PageSetup.Zoom = false;
                    xlWorkSheet.PageSetup.FitToPagesWide = 1;
                    xlWorkSheet.PageSetup.FitToPagesTall = 1;
                    xlWorkSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlPortrait;
                    xlWorkSheet.PageSetup.PaperSize = Microsoft.Office.Interop.Excel.XlPaperSize.xlPaperA4;
                    xlWorkSheet.Name = SetExcelSheetName(table, credGroup);
                    //Microsoft.Office.Interop.Excel.Range usedRange = xlWorkSheet.UsedRange;
                    //xlWorkSheet.Sort.SortFields.Add(xlWorkSheet.UsedRange.Columns["A"], Microsoft.Office.Interop.Excel.XlSortOn.xlSortOnValues, //Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending, System.Type.Missing, Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal);

                    // format cells
                    xlApp.Range["2:2"].Select();
                    xlApp.Application.Cells.EntireColumn.AutoFit();
                    xlApp.Application.Cells.EntireRow.AutoFit();
                    xlApp.Application.Cells.EntireColumn.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    xlApp.Application.Cells.EntireRow.HorizontalAlignment = XlHAlign.xlHAlignLeft;

                    xlApp.Application.Range["$A$2"].Select();

                    xlWorkBook.Password = ((string)ClientGroup.GetCredPass(credGroup));

                    if (Directory.Exists((string)ClientGroup.GetCredPath(credGroup)))
                    {
                        xlWorkBook.SaveAs(filePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        xlWorkBook.Close(true, misValue, misValue);
                        xlApp.Quit();
                    }
                    else
                    {
                        throw new Exception("Directory " + (string)ClientGroup.GetCredPath(credGroup) + " doesn't exist!");
                    }
                }
            }
            else // if client wants more than one sheet...
            {
                filePath = null;
                throw new Exception("Wrong Excel method called on this client group.");
            }
            return filePath;
        }

        private string BreakoutExcelOutput(List<System.Data.DataTable> tableList, ClientGroup credGroup) // for client groups with multiple sheets
        {
            string fileName = "";
            if (tableList.Count > 0 && ClientGroup.GetCredBreakout(credGroup) == "YES") // client wants multiple sheets regardless of how many files (each time method is called, it will create a new file)
            {
                // create excel file
                Console.WriteLine("Creating Excel file...");
                Microsoft.Office.Interop.Excel.Application xlApp;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                xlApp = new Microsoft.Office.Interop.Excel.Application();

                if (xlApp == null)
                {
                    throw new Exception("Excel is not properly installed!");
                }
                else
                { // setup parts of excel sheet that are the same for all groups
                    xlApp.Visible = false;
                    xlApp.DisplayAlerts = false;
                    xlApp.ScreenUpdating = false;
                    var xlWorkBooks = xlApp.Workbooks;
                    xlWorkBook = xlWorkBooks.Add(misValue);

                    foreach (System.Data.DataTable table in tableList) // loop through every table in the passed-in list to create separate sheets for the tables
                    {
                        xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Sheets[xlWorkBook.Sheets.Count]; // set current worksheet to end of sheets
                        //((Microsoft.Office.Interop.Excel.Worksheet)xlApp.ActiveWorkbook.Sheets[xlApp.ActiveWorkbook.Sheets.Count]).Select(); // selects first worksheet

                        // Excel column value formatting
                        xlWorkSheet.Columns["A:A"].NumberFormat = "@";
                        xlWorkSheet.Columns["J:J"].NumberFormat = "@";
                        xlWorkSheet.Columns["G:G"].NumberFormat = "mm/dd/yyyy";
                        xlWorkSheet.Columns["H:H"].NumberFormat = "mm/dd/yyyy";
                        xlWorkSheet.Columns["K:K"].NumberFormat = "mm/dd/yyyy";
                        xlWorkSheet.Columns["M:M"].NumberFormat = "mm/dd/yyyy";
                        xlWorkSheet.Columns["N:N"].NumberFormat = "mm/dd/yyyy";
                        xlWorkSheet.Columns["O:O"].NumberFormat = "mm/dd/yyyy";
                        xlWorkSheet.Columns["P:P"].NumberFormat = "mm/dd/yyyy";

                        // sets counts for rows and columns from current table from passed-in list
                        int rows = table.Rows.Count;
                        int columns = table.Columns.Count;

                        // Add the +1 to allow room for column headers
                        var data = new object[rows + 1, columns];

                        // populate headers in excel sheet
                        int colHeaderCt = 0; // default value for number of headers in the sheet

                        foreach (DataColumn col in table.Columns) // loop through every column in the passed-in table
                        {
                            data[0, colHeaderCt] = (string)col.ColumnName; // set the specified cell (column from loop, first row only) to the table's column name
                            colHeaderCt++;
                        }

                        // populate data object for inserting into excel from datatable
                        for (int row = 0; row < rows; row++)
                        {
                            for (int col = 0; col < columns; col++)
                            {
                                if ((string)table.Rows[row][col] == "01/01/0001" || (string)table.Rows[row][col] == "01/01/1000" || (string)table.Rows[row][col] == "1/1/1000" || (string)table.Rows[row][col] == "1/1/0001")  // if any of the cells are dates that are any of these placeholders...
                                {
                                    data[row + 1, col] = ""; // change the Excel date value to an empty string
                                }
                                else // if they're not any of these empty date placeholder values...
                                {
                                    data[row + 1, col] = table.Rows[row][col]; // add the value to the Excel object
                                }

                            }
                        }

                        // Write this data to the excel worksheet.
                        Range beginWrite = (Range)xlWorkSheet.Cells[1, 1];
                        Range endWrite = (Range)xlWorkSheet.Cells[rows + 1, columns];
                        Range sheetData = xlWorkSheet.Range[beginWrite, endWrite];
                        sheetData.Value2 = data;
                        sheetData.Select();

                        // page settings
                        xlApp.Cells.Font.Size = 11;
                        xlWorkSheet.PageSetup.Zoom = false;
                        xlWorkSheet.PageSetup.FitToPagesWide = 1;
                        xlWorkSheet.PageSetup.FitToPagesTall = 1;
                        xlWorkSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlPortrait;
                        xlWorkSheet.PageSetup.PaperSize = Microsoft.Office.Interop.Excel.XlPaperSize.xlPaperA4;
                        xlWorkSheet.Name = SetExcelSheetName(table, credGroup);
                        //Microsoft.Office.Interop.Excel.Range usedRange = xlWorkSheet.UsedRange;
                        //xlWorkSheet.Sort.SortFields.Add(xlWorkSheet.UsedRange.Columns["A"], Microsoft.Office.Interop.Excel.XlSortOn.xlSortOnValues, 
                        //Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending, System.Type.Missing, Microsoft.Office.Interop.Excel.XlSortDataOption.xlSortNormal);

                        // format cells
                        xlApp.Range["2:2"].Select();
                        xlApp.Application.Cells.EntireColumn.AutoFit();
                        xlApp.Application.Cells.EntireRow.AutoFit();
                        xlApp.Application.Cells.EntireColumn.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                        xlApp.Application.Cells.EntireRow.HorizontalAlignment = XlHAlign.xlHAlignLeft;

                        // add new sheet
                        if (xlWorkBook.Sheets.Count < tableList.Count) // don't add more sheets than there are tables -> if there are still more tables, add a sheet
                        {
                            xlWorkBook.Worksheets.Add(After: xlWorkBook.Sheets[xlWorkBook.Sheets.Count]);
                        }
                    }

                    // Set Excel file name
                    fileName = SetExcelFileName(tableList[tableList.Count - 1], credGroup); // call function to set file name based on table name/return type

                    filePath = (ClientGroup.GetCredPath(credGroup) + fileName + ".xls"); // set full file path with file name
                    int count = 0; // dupe count default

                    while ((File.Exists(filePath))) // while the file path already exists...
                    {
                        fileName = (fileName + "_" + count); // update the file name with a dupe count value
                        filePath = (ClientGroup.GetCredPath(credGroup) + fileName + ".xls"); // update the file path with the new file name
                        count++;
                    }

                    // Select the first cell in the worksheet.
                    xlApp.Application.Range["$A$2"].Select();
                    xlWorkBook.Password = ((string)ClientGroup.GetCredPass(credGroup));

                    if (Directory.Exists((string)ClientGroup.GetCredPath(credGroup)))
                    {
                        xlWorkBook.SaveAs(filePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                        xlWorkBook.Close(true, misValue, misValue);
                        xlApp.Quit();
                    }
                    else
                    {
                        throw new Exception("Directory " + (string)ClientGroup.GetCredPath(credGroup) + " doesn't exist!");
                    }
                }
            }
            else // if client does not want multiple Excel sheets...
            {
                filePath = null;
                throw new Exception("Wrong Excel method called on this client group or empty list of datatables.");
            }
            return filePath;
        }

        private void SetColumns(System.Data.DataTable fullTable, System.Data.DataTable emptyTable) // sets columns for specified data table from another table's columns
        {
            emptyTable.Columns.Clear(); // clears out any existing columns from the table that is supposed to be updated

            foreach (DataColumn col in fullTable.Columns) // add correct number of columns to output data table from copied table
            {
                DataColumn bColumn;
                bColumn = new DataColumn();
                bColumn.ColumnName = col.ColumnName;
                emptyTable.Columns.Add(bColumn);
            }
        }

        private List<Dictionary<int, ReturnAccount>> CustomReturnSplitting(ClientGroup creditorGroup, bool distinct) // This method is only called if already confirmed that the client wants custom splitting according to the DB
        {
            List<Dictionary<int, ReturnAccount>> customSplitDicts = new List<Dictionary<int, ReturnAccount>>(); // create temp list of dictionaries for custom accounts

            if (ClientGroup.GetCredSplitStatus(creditorGroup) != null && ClientGroup.GetCredSplitStatus(creditorGroup)[0] != "") // if status is not empty/null... TODO: change in V2
            {
                foreach (string splitStatus in ClientGroup.GetCredSplitStatus(creditorGroup)) // loops through array of custom split statuses listed for client in DB
                {
                    if (StatusValid(splitStatus)) // calls function to do status code validation using data table of bloodhound statuses
                    {
                        Dictionary<int, ReturnAccount> customSplitDict = new Dictionary<int, ReturnAccount>(); // create new dictionary for list of custom dicts

                        foreach (ReturnAccount account in allReturnsInGroup.Values.ToList()) // loops through every account currently in the list of all returns
                        {
                            if ((string)ReturnAccount.GetReturnStatus(account) == (string)splitStatus) // if the current account's status = the custom split status...
                            {
                                SwitchReturnLists(account, allReturnsInGroup, customSplitDict, distinct); // switch the account over to the new dictionary
                            }
                        }
                        customSplitDicts.Add(customSplitDict); // add the dictionary to the final list that will be returned
                    }
                    else // if the specified split status cannot be validated...
                    {
                        throw new Exception("Status " + splitStatus + " for client group " + creditorGroup + " is invalid.");
                    }
                }
            }
            return customSplitDicts;
        }

        private void SwitchReturnLists(ReturnAccount item, Dictionary<int, ReturnAccount> original, Dictionary<int, ReturnAccount> newDict, bool distinctItem) // switches or copies a return account/dictionary value from one list to another
        {
            int dupeKeyCounter = 0; // default value for duplicate key counter

            if (original.ContainsValue(item)) // if the original dictionary contains the account to be switched/copied...
            {
                var keysWithMatchingValues = original.Where(p => p.Value == item).Select(p => p.Key); // store the key(s) where the value matches the parameter account

                foreach (var key in keysWithMatchingValues.ToList()) // loops through the list of keys of matching pairs
                {
                    if (key.ToString().Length > 0) // if the key is not empty...
                    {
                        if (distinctItem) // if the client wants accounts separated into bankos/settleds/etc. removed from the returns file...
                        {
                            int copyKey = key;
                            newDict.Add(copyKey, item); // add account to new dictionary
                            original.Remove(copyKey); // remove account from old dictionary
                        }
                        else // if the client wants accounts separated into bankos/settleds/etc. to also stay in the returns file...
                        {
                            int copyKey = key;
                            newDict.Add(key, item); // add account to new dictionary
                        }

                        dupeKeyCounter++;
                    }
                }
                if (dupeKeyCounter > 1) // if whole row of SQL table is the same as another...
                {
                    throw new Exception("Duplicate BloodHound return data error!");
                }
            }
            else // if the original dictionary doesn't contain the account in question...
            {
                throw new Exception("Incorrect account or dictionary passed in!");
            }
        }

        private System.Data.DataTable ReturnsListToTable(Dictionary<int, ReturnAccount> inDict, string parameterTableName, ClientGroup creditorGroup) // takes a list of return accounts and turns it into a data table that can be looped through to print to an excel sheet
        {
            if (inDict.Values.Count > 0 && inDict != null) // if passed in dictionary is not empty...
            {
                bankoTableNames.Clear(); // clear out list from last time function was called

                // these table/file names are bankruptcies
                bankoTableNames.Add("Bankos");
                bankoTableNames.Add("Banko");
                bankoTableNames.Add("Bank");
                bankoTableNames.Add("Bankruptcy");
                bankoTableNames.Add("Bankruptcies");

                // these columns should NEVER be included in the table
                List<string> colDump = new List<string>();
                colDump.Add("Status");
                colDump.Add("Creditor");
                colDump.Add("Debtor Number");

                // these are the banko-specific columns that are included in ALL breakouts and ONLY banko files for itemized files
                List<string> bankoColDump = new List<string>();
                bankoColDump.Add("Chpt / Case Number");
                bankoColDump.Add("Filing Date");
                bankoColDump.Add("Discharge Date");
                bankoColDump.Add("Proof Date");
                bankoColDump.Add("Dismissal Date");

                // create new table to output and later output to excel
                System.Data.DataTable newTable = new System.Data.DataTable();

                // fill new table to return at the end of the function and add columns
                foreach (DataColumn col in bhDt.Columns)
                {
                    // if the column name is in colDump it should be skipped; otherwise continue
                    if (!colDump.Contains(col.ColumnName))
                    {
                        // if the column IS a banko column it needs to meet one of these requirements
                        if (bankoColDump.Contains(col.ColumnName))
                        {
                            // if the files are NOT itemized OR the table is for bankruptcies then the banko columns CAN be included in the new data table
                            if (ClientGroup.GetCredItemized(creditorGroup) == "NO" || bankoTableNames.Contains(parameterTableName))
                            {
                                DataColumn Column;
                                Column = new DataColumn();
                                Column.ColumnName = col.ColumnName;
                                newTable.Columns.Add(Column);
                            }
                        }

                        // if the current column in the loop is NOT in the colDump list and NOT a banko then we add it to the new table
                        else
                        {
                            DataColumn Column;
                            Column = new DataColumn();
                            Column.ColumnName = col.ColumnName;
                            newTable.Columns.Add(Column);
                        }
                    }
                }

                // create a new row for each account in the data table that will be returned
                foreach (ReturnAccount returnAcct in inDict.Values.ToList())
                {
                    // regular columns:
                    DataRow row = newTable.NewRow(); // create a fresh new row with the following data:
                    row[0] = (string)ReturnAccount.GetReturnAcctNumber(returnAcct);
                    row[1] = (string)ReturnAccount.GetReturnGuarantor(returnAcct);
                    row[2] = (string)ReturnAccount.GetReturnPatient(returnAcct);
                    row[3] = (string)ReturnAccount.GetReturnDescription(returnAcct);
                    row[4] = (string)ReturnAccount.GetReturnAmount(returnAcct);
                    row[5] = (string)ReturnAccount.GetReturnAmtListed(returnAcct);
                    row[6] = ReturnAccount.FormatDateToString((DateTime)(ReturnAccount.GetReturnDate(returnAcct)));
                    row[7] = ReturnAccount.FormatDateToString((DateTime)(ReturnAccount.GetReturnLastPayDate(returnAcct)));
                    row[8] = (string)ReturnAccount.GetReturnLastPayAmount(returnAcct);
                    row[9] = (string)ReturnAccount.GetReturnCredNumber(returnAcct);
                    row[10] = ReturnAccount.FormatDateToString((DateTime)(ReturnAccount.GetReturnServiceDate(returnAcct)));

                    // banko columns:
                    if (!MissingBankoData(newTable)) // if the table is not missing banko columns
                    {
                        row[11] = (string)ReturnAccount.GetReturnBhCode(returnAcct);
                        row[12] = ReturnAccount.FormatDateToString((DateTime)(ReturnAccount.GetReturnBhFileDate(returnAcct)));
                        row[13] = (string)ReturnAccount.GetReturnBhDischarge(returnAcct);
                        row[14] = (string)ReturnAccount.GetReturnBhProof(returnAcct);
                        row[15] = ReturnAccount.FormatDateToString((DateTime)(ReturnAccount.GetReturnBhDismiss(returnAcct)));
                    }
                    newTable.Rows.Add(row); // add the row to the data table
                }
                newTable.TableName = parameterTableName; // important: this is used to determine the name of the sheets and files sent to clients' systems
                return newTable;
            }
            else // if dictionary passed in is empty...
            {
                return null;
            }
        }

        private string SetExcelSheetName(System.Data.DataTable table, ClientGroup creditorGroup) // sets the name of an excel sheet with the table name
        {
            string sheetName = "";
            if (table.TableName != "" && table.TableName != null)
            {
                if (ClientGroup.GetCredBreakout(creditorGroup) == "NO" && ClientGroup.GetCredItemized(creditorGroup) == "YES" || ClientGroup.GetCredSplitType(creditorGroup) == "ITEMIZED") // if the client group does not want multiple sheets but wants multiple files OR has an itemized custom split...
                {
                    switch (table.TableName)
                    {
                        case "Bankruptcy":
                            sheetName = "Banko";
                            break;
                        case "Settled in Full":
                            sheetName = "Settled";
                            break;
                        case "Returns":
                            sheetName = "Data";
                            break;
                        default:
                            sheetName = "Data";
                            break;
                    }
                }

                else if (ClientGroup.GetCredBreakout(creditorGroup) == "YES" || ClientGroup.GetCredSplitType(creditorGroup) == "BREAKOUT") // if the client group wants multiple sheets OR wants a custom split in a separate sheet...
                {
                    if (CreditorNumberValid(table.TableName)) // if the table name is a valid creditor number in bloodhound...
                    {
                        sheetName = table.TableName + "_Returns"; // the file is a breakout/split by creditor numbers and this sheet is not custom; use default format
                    }
                    else // if the table name is not a valid creditor number...
                    {
                        sheetName = table.TableName; // it is a custom/differing, breakout split; the sheet can be named after the table name
                    }
                }
                else // client doesn't want breakout OR itemization; set sheet name to Data
                {
                    sheetName = "Data";
                }
            }
            return sheetName;
        }

        public string SetExcelFileName(System.Data.DataTable table, ClientGroup creditorGroup) // set name of excel file from itemization value and passed-in table name
        {
            string fileName = "";
            string useThisDateFormat = "";
            string dateFormatWeeklyClients = "yyyy-MM-dd";
            string dateFormatMonthlyClients = "MM-dd-yyyy";

            // file name date formatting is different for weekly and monthly clients
            if (ClientGroup.GetCredFrequency(creditorGroup) == "MONTHLY")
            {
                useThisDateFormat = dateFormatMonthlyClients;
            }
            else
            {
                useThisDateFormat = dateFormatWeeklyClients;
            }

            // set file name using the correct date format based on whether client wants multiple files or not
            if (ClientGroup.GetCredItemized(creditorGroup) == "NO") 
            {
                // if the creditor group does not want multiple files...
                fileName = ((ClientGroup.GetCredName(creditorGroup) + " - Returns - " + firstDate.ToString(useThisDateFormat) + " - " + secondDate.ToString(useThisDateFormat)).ToString()); // set to the default format for single-file returns
            }
            else if (ClientGroup.GetCredItemized(creditorGroup) == "YES") 
            {
                // if creditor group wants split/multiple files...
                if (table != null && table.TableName != null && creditorGroup != null) 
                {
                    // if no important variables are null/empty...
                    fileName = (ClientGroup.GetCredName(creditorGroup) + " - " + table.TableName + " - " + firstDate.ToString(useThisDateFormat) + " - " + secondDate.ToString(useThisDateFormat)).ToString(); 
                    // set format for file name based on the table name/return type (E.g. Bankruptcy, Returns)
                }
            }

            // if client's specified option for itemization is not YES or NO...
            else
            {
                fileName = null;
                throw new Exception("Parameter(s) null or missing in SetExcelFileName function or invalid itemization value!");
            }

            //TODO: remove after testing
            return "TEST_" + fileName;
        }

        private void Email(string returnTp, string file, ClientGroup clientGroup) // calls function to create excel files and then emails them out
        {
            Console.WriteLine("Sending email...");

            // for addresses that should be CCed
            List<string> toCC = new List<string>();

            // ALWAYS CC REPORTS!
            toCC.Add("reports@americollect.com".ToUpper()); 
            
            string address = ""; // main address that email will be sent to


            if (returnTp.Trim().ToUpper() == "RETURN" || returnTp.Trim().ToUpper() == "RETURNS")
            {
                address = ClientGroup.GetCredReturnEmails(clientGroup)[0]; // first in array = main address
                for (int i = 1; i < ClientGroup.GetCredReturnEmails(clientGroup).Length; i++)
                {
                    toCC.Add((string)ClientGroup.GetCredReturnEmails(clientGroup)[i]); // others: CC
                }
            }
            else if (returnTp.Trim().ToUpper() == "BANKRUPTCY")
            {
                address = ClientGroup.GetCredBankoEmails(clientGroup)[0]; // first in array = main address
                for (int i = 1; i < ClientGroup.GetCredBankoEmails(clientGroup).Length; i++)
                {
                    toCC.Add((string)ClientGroup.GetCredBankoEmails(clientGroup)[i]); // others: CC
                }
            }
            else if (returnTp.Trim().ToUpper() == "SETTLED IN FULL")
            {
                address = ClientGroup.GetCredSettleEmails(clientGroup)[0]; // first in array = main address
                for (int i = 1; i < ClientGroup.GetCredSettleEmails(clientGroup).Length; i++)
                {
                    toCC.Add((string)ClientGroup.GetCredSettleEmails(clientGroup)[i]); // others: CC
                }
            }
            else // use returns address if all else fails so the client gets their returns
            {
                address = ClientGroup.GetCredReturnEmails(clientGroup)[0]; // first in array = main address
                for (int i = 1; i < ClientGroup.GetCredReturnEmails(clientGroup).Length; i++)
                {
                    toCC.Add((string)ClientGroup.GetCredReturnEmails(clientGroup)[i]); // others: CC
                }
            }

            address = "ashleys@americollect.com".ToUpper(); // TODO: FOR TESTING --> REMOVE

            if (address != "") // if a main email address was set during the switch statement...
            {
                StringBuilder sb = new StringBuilder(); // string builder for emails listed in body of email
                sb.Append("<br>" + address);

                if (toCC.Count > 0)
                {
                    foreach (string ccAddy in toCC.Distinct()) // loop through each unique email address in list
                    {
                        sb.Append("<br>" + ccAddy);
                    }
                }
                else
                {
                    sb.Append("");
                }

                // TODO: TESTING - Remove after
                toCC.Add("kellyR@americollect.com".ToUpper());
                toCC.Add("josephg@americollect.com".ToUpper());

                toCC.RemoveAll(x => x != "KELLYR@AMERICOLLECT.COM" && x != "JOSEPHG@AMERICOLLECT.COM" && x != "REPORTS@AMERICOLLECT.COM");

                string subject = file; // sets subject to correct file name

                string body = "Hello:<br>Attached is your return report.<br>Let us know if you have any questions.<br>Americollect Support Team 800 - 838 - 0100<br><br>THIS IS A TEST THAT WOULD HAVE BEEN SENT TO: " + sb.ToString() + "<br><br>Please contact Amy Cerkas, Adam Rathsack, or Joe Gramling if this information is incorrect.<br>Thank you!";

                using (MailMessage message = new MailMessage("macro@americollect.com", address)) // FROM address
                using (SmtpClient client = new SmtpClient("mail.americollect.com")) // SMTP client
                {
                    Attachment report = new Attachment(file, System.Net.Mime.MediaTypeNames.Application.Octet); // Excel file attached
                    message.Subject = subject;
                    message.Body = body;
                    message.IsBodyHtml = true;
                    message.Attachments.Add(report);
                    message.Priority = MailPriority.High;
                    
                    foreach (string ccAddress in toCC.Distinct()) // loops through all emails in the CC list
                    {
                        message.CC.Add(ccAddress); // adds address to the CC field in email
                        emailedToList.Add(ccAddress); // adds emails to list of emails sent to for log records table in DB
                    }

                    client.Port = 25;
                    client.Credentials = CredentialCache.DefaultNetworkCredentials;
                    client.EnableSsl = false;

                    client.Send(message);
                    report.Dispose();

                    emailedToList.Add(address); // adds emails to list of emails sent to for log records table in DB
                    totalFilesEmailed++; // updates count for total number of sent emails for log records table in DB
                }
            }
        }

        private void FilterOutSharedAccounts() // shared accounts don't necessarily have LA2 status anymore so we need to check for missing client account #s
        {
            if (bhDt.Rows.Count > 0) // if data was pulled and stored from bloodhound...
            {
                for (int dtRow = 0; dtRow < bhDt.Rows.Count; dtRow++) // loop through all rows in data table of all bloodhound records for creditor group
                {
                    if (bhDt.Rows[dtRow]["Your Account #"].ToString() == "" || bhDt.Rows[dtRow]["Your Account #"] == null || bhDt.Rows[dtRow]["Linked"].ToString() != "") // if client account # is empty or a linked account is listed...
                    {
                        bhDt.Rows.RemoveAt(dtRow); // remove the row with the missing account number from the data table containing all records for cred group
                    }
                }

                // remove the linked account column after removing all rows that have data for it
                bhDt.Columns.Remove("Linked");
            }
        }

        private System.Data.DataTable GetLogDailyRecords() // fetches all log records from SQL and stores them in a temporary data table
        {
            System.Data.DataTable tempTable = new System.Data.DataTable(); // creates temporary data table for log records

            using (OdbcConnection cnSQL = new OdbcConnection(sqlConnStr))
            {
                cnSQL.Open();
                selectSQL = "SELECT * FROM `RunLogs` WHERE `logId` >= 0 AND `runAsDate` = ?"; // grabs all run log rows that have a valid ID and were run as today

                using (OdbcCommand selectCMD = new OdbcCommand(selectSQL, cnSQL))
                {
                    selectCMD.Parameters.Add("@runAsDate", OdbcType.Date).Value = ((DateTime)curDt).ToShortDateString(); // program's current date (adjusted)

                    OdbcDataAdapter da = new OdbcDataAdapter(selectCMD);

                    if (da.ToString().Length > 0) // if any data is found...
                    {
                        da.Fill(tempTable);  // adds all found data to temp table
                    }
                }
            }

            // Fix blank cells to prevent errors later
            if (tempTable.Rows.Count > 0)
            {
                for (int b = 0; b < tempTable.Rows.Count; b++)
                {
                    for (int a = 0; a < tempTable.Columns.Count; a++)
                    {
                        if (tempTable.Rows[b][a] == System.DBNull.Value || tempTable.Rows[b][a] == null || tempTable.Rows[b][a].ToString() == "") // stores null values as empty strings to prevent errors later
                        {
                            tempTable.Rows[b][a] = "";
                        }
                    }
                }
            }
            return tempTable;
        }

        private void CreateReturnSummary(List<string> frequencyTypes) // generates a brief PDF report of returns sent
        {
            // set name for file, filepath, and handle duplicate files
            string reportName = "";
            string reportFilePath = "";

            if (frequencyTypes.Count > 0) // if frequencies exist in the parameter list...
            {
                foreach (string freq in frequencyTypes) // loops through the frequencies listed to run today, creating a separate file for each
                {
                    StringBuilder reportSb; // string that will be converted to the text in the pdf file
                    reportName = (string)curDt.Date.ToString("MMddyyyy") + "_" + freq + "_UNIVERSAL"; // set file name with run date and frequency
                    reportFilePath = (string)@"G:\Clients\Universal Return Macro\" + reportName + ".pdf"; // update path with new file name

                    // handle duplicate/existing files
                    int dupeCt = 1;
                    while (File.Exists(reportFilePath))
                    {
                        reportName = reportName + "_" + dupeCt;
                        reportFilePath = (string)@"G:\Clients\Universal Return Macro\" + reportName + ".pdf"; // update path with new file name
                        dupeCt++;
                    }

                    // stream for saving file to specific location
                    FileStream fileStream = File.Create(reportFilePath);

                    // create PDF file
                    Document document = new Document(PageSize.A4, 25, 25, 30, 30);
                    PdfWriter writer = PdfWriter.GetInstance(document, fileStream);
                    document.AddTitle(reportName);
                    document.Open(); // open the document to enable you to write to the document

                    if (emailedToList.Count > 0) // if any emails were sent out for this frequency...
                    {
                        // create stringbuilder for text that should be put in pdf document
                        reportSb = new StringBuilder();
                        var nL = Environment.NewLine; // create a new line in the string
                        reportSb.Append("Report Date: " + DateTime.Now.ToString("MM/dd/yyyy") + nL + "Program Run Date: " + curDt.ToString("MM/dd/yyyy") + nL + "Client Group Frequency: " + freq.ToString() + nL + nL); // header with dates and frequency type

                        reportSb.Append("Client Groups Processed: " + nL);
                        foreach (string clientEmailed in passFailClientRecordList) // body: loops through all client groups processed and adds to sb
                        {
                            reportSb.Append(clientEmailed + nL);
                        }

                        reportSb.Append(nL + "Receiving Addresses: " + nL);
                        foreach (string emailAddy in emailedToList.Distinct()) // body: loops through all email addresses sent files and adds to sb
                        {
                            reportSb.Append(emailAddy + nL);
                        }

                        reportSb.Append(nL + nL + "Number of Files Created: " + totalFilesEmailed);
                    }
                    else // if no emails were sent with return files...
                    {
                        // create stringbuilder with text to put in document
                        reportSb = new StringBuilder();
                        var nL = Environment.NewLine; // create a new line in the string
                        reportSb.Append("Report Date: " + DateTime.Now.ToString("MM/dd/yyyy") + nL);
                        reportSb.Append("No return files were sent for the run date: " + curDt.ToString("MM/dd/yyyy"));
                    }

                    // Add a simple and wellknown phrase to the document in a flow layout manner
                    document.Add(new Paragraph(reportSb.ToString()));

                    // close the document
                    document.Close();

                    // close the writer instance
                    writer.Close();

                    // always close open filehandles explicity
                    fileStream.Close();

                    // email the file
                    string subject = reportName; // sets subject to file name
                    string body = "Hello:<br>Attached is your Universal Returns summary report.";
                    string mainAddress = "josephg@Americollect.com"; // for testing - TODO: CHANGE
                    string[] ccReport = { }; // array of addresses to CC report email in

                    using (MailMessage message = new MailMessage("macro@americollect.com", mainAddress)) // FROM address, TO address
                    using (SmtpClient client = new SmtpClient("mail.americollect.com")) // SMTP client
                    {
                        Attachment report = new Attachment(reportFilePath, System.Net.Mime.MediaTypeNames.Application.Octet); // Excel file attached
                        message.Subject = subject;
                        message.Body = body;
                        message.IsBodyHtml = true;
                        message.Attachments.Add(report);
                        message.Priority = MailPriority.High;

                        if (ccReport.Length > 0)
                        {
                            foreach (string ccAddress in ccReport.Distinct()) // loops through all emails in the CC list
                            {
                                message.CC.Add(ccAddress); // adds address to the CC field in email
                                emailedToList.Add(ccAddress); // adds emails to list of emails sent to for log records table in DB
                            }
                        }

                        client.Port = 25;
                        client.Credentials = CredentialCache.DefaultNetworkCredentials;
                        client.EnableSsl = false;

                        client.Send(message);
                        report.Dispose();
                    }
                }
            }
            else // if there are no frequencies in list...
            {
                reportFilePath = ""; // make file path empty so a file isn't sent
            }
        }

        private bool MissingBankoData(System.Data.DataTable table) // returns true if banko columns are missing from passed-in table
        {
            // array of banko column names is global!

            // number of banko-specific columns actually in the provided table
            int bankoColsCount = 0;

            // loop through all columns in the table to determine if they are banko columns
            foreach (DataColumn col in table.Columns)
            {
                // if the column name is on the banko list then the table already has banko columns
                if (bankoColNames.Contains(col.ColumnName))
                {
                    bankoColsCount++;
                }
            }

            // if the actual number of banko columns in the table matches the number of banko col names then the table has all the banko columns
            if (bankoColsCount == bankoColNames.Length)
            {
                return false;
            }

            // if the table has less than the number of banko column names then it is missing some or all banko columns
            else if (bankoColsCount < bankoColNames.Length)
            {
                return true;
            }

            // if the table has more than the number of banko column names than something is very wrong :(
            else
            {
                return true; // we need to return something for this function to not error but it shouldn't matter since we're throwing an exception next
                throw new Exception("Data table has more bankruptcy columns than expected!");
            }
        }

        public void WriteLogFile(bool isSuccess, string prob = "")
        {
            string fullName = System.Reflection.Assembly.GetEntryAssembly().Location;         //path and name
            string myName = Path.GetFileNameWithoutExtension(fullName);     //just name
            string logPath;
            string errMessage;

            //determine if we're writing a success or a fail file
            if (isSuccess)
            {
                logPath = @"\\criticalprocess\Automation\Log\Success\";
                errMessage = "Great job!";
            }
            else
            {
                logPath = @"\\criticalprocess\Automation\Log\Errors\";
                errMessage = "Ya blew it!\n" + prob;
            }

            //create text file
            using (StreamWriter writer = new StreamWriter(logPath + DateTime.Now.ToString("yyyyMMdd_HH.mm.ss~~") + myName + ".exe~~" + DateTime.Now.ToString("dddd").ToUpper() + ".txt", true))
            {
                writer.WriteLine(errMessage);
            }
        }

        private void ErrorEmailer(string inErr)
        {
            string processLoc = System.Reflection.Assembly.GetEntryAssembly().Location;                 //path and name
            string processName = System.IO.Path.GetFileNameWithoutExtension(processLoc);             //just name

            System.Net.Mail.MailMessage msg = new System.Net.Mail.MailMessage();
            msg.To.Add("josephg@americollect.com");

            msg.From = new System.Net.Mail.MailAddress("macro@americollect.com");
            msg.Subject = processName.ToUpper() + " ERRORED!_" + DateTime.Now.ToString("yyyy-MM-dd");
            msg.Body = "<font face=Calibri>This automated process encountered an error and may need to be reran.<br><br>" +
            "Process Name: " + processName + "<br>Process Location: " + processLoc + "<br><br>Error Description: " + inErr + "<br>" +
            "Date of Error: " + DateTime.Now.ToString("MM/dd/yyyy") + "<br>Time of Error: " + DateTime.Now.ToString("HH.mm.ss") +
            "<br>Day of Error: " + DateTime.Now.ToString("dddd").ToUpper() + "<br><br></font>";
            msg.IsBodyHtml = true;
            System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient("mail.americollect.com");
            smtp.Send(msg);
        }
        #endregion
    }
}