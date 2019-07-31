using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Data;
using DatabaseAccess;
using System.Data.SqlClient;

namespace SheetsQuickstart
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "Google Sheets API .NET Test App";
        static List<CustomerRoster> customerRecordsList = new List<CustomerRoster>();
        static IList<IList<Object>> values;

        static void Main(string[] args)
        {            
            UserCredential credential;

            using (var stream =
                new FileStream("GoogleCredentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define request parameters.
            String spreadsheetId = "1hobYZgYiONQ4CiwkIdOk_ZFXaqwIguYF2294ChTWIVU";
            String range = "PY45 ACTIVE!A4:V";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);

            // Prints the names and majors of students in a sample spreadsheet:
            // https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
            ValueRange response = request.Execute();
            values = response.Values;

            FillRosterTable(); // where the magic happens.
        }

        private static void FillRosterTable()
        {
            const int PYStartNumber = 1;
            const int PY_ID_CONST = 45;
            if (values != null && values.Count > 0)
            {
                Console.WriteLine("PY RosterID, LastName, FirstName");
                foreach (var row in values)
                {
                    // Print columns A and E, which correspond to indices 0 and 4.
                    if (row.Count > 3)
                    {
                        Console.WriteLine("{0}, {1}, {2}", row[1], row[2], row[3]);
                        //CustomerRoster tempCust = new CustomerRoster();
                        if (Int32.TryParse(row[1].ToString(), out int tempRecordID))
                        {
                            if (tempRecordID < PYStartNumber)
                                continue;
                        }
                        else
                        {
                            Console.Error.WriteLine($"Could not parse rosterID!\n\nRecord Number: {row[1]}");
                        }

                        if (Int32.TryParse((row.Count > 11 ? row[11] : "").ToString(), out int tempISIS_ID))
                        {
                            // do something?
                        }
                        else
                        {
                            Console.Error.WriteLine($"Could not parse rosterID!\n\nRecord Number: {row[1]}");
                        } // end ISIS_ID parse.

                        if (DateTime.TryParse((row.Count > 4 ? row[4] : "").ToString(), out DateTime tempDOB))
                        {
                            // do something?
                        }
                        else
                        {
                            Console.Error.WriteLine($"Could not parse DOB!\n\nRecord Number: {row[1]}");
                        } // end DOBParse

                        if (DateTime.TryParse((row.Count > 5 ? row[5] : "").ToString(), out DateTime tempDOS))
                        {
                            // do something?
                        }
                        else
                        {
                            Console.Error.WriteLine($"Could not parse DOS!\n\nRecord Number: {row[1]}");
                        } // end DateOfService Parse.

                        if (DateTime.TryParse((row.Count > 10 ? row[10] : "").ToString(), out DateTime tempSubmissionDate))
                        {
                            //tempCust.RosterID = tempInt;
                        }
                        else
                        {
                            Console.Error.WriteLine($"Could not parse SubmissionDate!\n\nRecord Number: {row[1]}");
                        } // end SubmissionDate parse.

                        if (DateTime.TryParse((row.Count > 13 ? row[13] : "").ToString(), out DateTime tempIntakeDate))
                        {
                            //tempCust.RosterID = tempInt;
                        }
                        else
                        {
                            Console.Error.WriteLine($"Could not parse IntakeDate!\n\nRecord Number: {row[1]}");
                        } // end IntakeDate parse.

                        if (DateTime.TryParse((row.Count > 15 ? row[15] : "").ToString(), out DateTime tempPSAExpDate))
                        {
                            //tempCust.RosterID = tempInt;
                        }
                        else
                        {
                            Console.Error.WriteLine($"Could not parse PSAExpDate!\n\nRecord Number: {row[1]}");
                            tempPSAExpDate = DateTime.Parse("1/1/1777");
                        } // end PSAExpDate parse.

                        customerRecordsList.Add(new CustomerRoster
                        {
                            RosterID = tempRecordID,
                            LastName = (string)row[2],
                            FirstName = (string)row[3],
                            DOB = tempDOB,
                            DateOfService = tempDOS,
                            Staff = row.Count > 6 ? (string)row[6] : null,
                            EnrollmentType = row.Count > 7 ? (string)row[7] : null,
                            ReferredBy = row.Count > 8 ? (string)row[8] : null,
                            ReasonForVisit = row.Count > 9 ? (string)row[9] : null,
                            SubmissionDate = tempSubmissionDate,
                            ISIS_ID = tempISIS_ID,
                            SelfCertified = row.Count > 12 ? (string)row[12] : null,
                            IntakeDate = tempIntakeDate,
                            AgeGroup = row.Count > 14 ? (string)row[14] : null,
                            PSAExpDate = tempPSAExpDate,
                            YouthSchool = row.Count > 16 ? (string)row[16] : null,
                            Notes = row.Count > 19 ? (string)row[17] + row[18] + row[19] : null,
                            PhoneNumber = row.Count > 20 ? (string)row[20] : null,
                            Email = row.Count > 21 ? (string)row[21] : null,
                            PY_ID = PY_ID_CONST
                        });
                    }
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }

            SqlConnection conn = new SqlConnection(ConnectionAccess.ListingsString);
            conn.Open();
            foreach (CustomerRoster cust in customerRecordsList)
            {
                try
                {
                    CustomerRoster.addCustomerRosterBatch(conn, cust);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("Error adding customer to the database!");
                    Console.Error.WriteLine($"Customer Record#: {cust.RosterID}\nCustomerName: {cust.FirstName + " " + cust.LastName}\n");
                    Console.Error.WriteLine($"Exception: {ex}\n\n");
                }
            }
            Console.WriteLine("Roster Records imported. Press Any Key To Continue...");
            conn.Close();
            Console.Read();
        }
        
    }
}