using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.ServiceModel.Description;
using System.Text;
using System.Threading.Tasks;

namespace Ciaran.PublishChecker
{
   

    class Program
    { 
       private static IOrganizationService service;
        protected static string username;
        protected static string password;
        static string _ClientVersionStamp; //last time query for changes was completed.
        public static string URL;

        static string input; 

        static void Main(string[] args)
        {
            Console.WriteLine("MetaData changes Checker:");
            Listen();
        }





        #region HelperMethods
        public static void Listen()
        {
            input = Console.ReadLine();

            switch (input.ToLower())
            {
                case "exit":
                    break;
                case "help":
                    PrintHelp();
                    Listen();
                    break;
                case "init":
                    Initialize();
                    Listen();
                    break;
                case "initquery":
                    InitialQuery();
                    Listen();
                    break;
                case "queryall":
                    QueryAllEntities();
                    Listen();
                    break;
                default:
                    Console.WriteLine("Command not recognised");
                    Listen();
                    break;
            }
        }

        public static void Initialize()
        {
            Console.WriteLine("Initializing...");
            Console.WriteLine("Please enter your username");
            username = Console.ReadLine();

            Console.WriteLine("Please enter your password");
            password = ReadPassword();

            Console.WriteLine("Please enter the target CRM: (eg. https://crm.crm.ie)");
            URL = Console.ReadLine();

            service = getOrgService();
            Console.WriteLine("Initialized");

        }

        public static void PrintHelp()
        {
            Console.WriteLine("Application commands: \n 1.'initquery' - run the first query to have a benchmark metadata timestamp. \n 2.'queryall' - Query for changes after the initial timestamp. \n 3. 'exit' - close the application. \n 4. 'help' - display all commands in the application. \n 5. 'init' - initialize the application, input your credentials.");
        }
        private static  IOrganizationService getOrgService()
        {
            //retreivs org service, returns false if fails. 
            //Get credentials from config file
            //Authenticate
            //Get and return org service 
            string crmurl = URL; 
            string crmuser = username;
            string crmpass = password;
            String apiuri = "/XRMServices/2011/Organization.svc";
            IOrganizationService _service;

            Uri organizationUrl = new Uri(crmurl + apiuri);
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            try
            {
                ClientCredentials credentials = new ClientCredentials();
                credentials.UserName.UserName = crmuser;
                credentials.UserName.Password = crmpass;

                var organizationService = (IOrganizationService)new OrganizationServiceProxy(organizationUrl, null, credentials, null);
                return organizationService;

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while connecting to CRM: " + ex.Message);
                return null; 
            }
        }

        public static void InitialQuery()
        {
            var spinner = new Spinner(5,0);

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Query Started...");
            spinner.Start();
            RetrieveMetadataChangesRequest req = new RetrieveMetadataChangesRequest()
            {
                ClientVersionStamp = null
            };
            var response = (RetrieveMetadataChangesResponse)service.Execute(req);
            spinner.Stop();
            Console.WriteLine("\n Timestamp Response: " + response.ServerVersionStamp);
            Console.ForegroundColor = ConsoleColor.White;

            if (!log(response.ServerVersionStamp, "TimeStamp", false, true))
            {
                Console.WriteLine("Error storing timestamp.");
            }
            else Console.WriteLine("Timestamp stored.");
        }

        public static void QueryAllEntities()
        {
            string _timestamp = "";
            int DaysSinceModified = 0;

            var spinner = new Spinner(5, 0);
            Console.ForegroundColor = ConsoleColor.Green;
            spinner.Start();


            EntityMetadata md;

            if (_ClientVersionStamp != null)
            {
                _timestamp = _ClientVersionStamp;

            }else
            {
                try
                {
                    _timestamp = GetStoredTimestamp();
                    _ClientVersionStamp = _timestamp;
                    Console.WriteLine("Stored timestamp found: " + _timestamp );
                }catch(Exception e)
                {
                    Console.WriteLine("Error getting stored timestamp: " + e.Message);
                }
            }

            var EntityFilter = new MetadataFilterExpression(LogicalOperator.And);
            //EntityFilter.Conditions.Add(new MetadataConditionExpression("SchemaName", MetadataConditionOperator.Equals, "ServiceAppointment"));
            var entityQueryExpression = new EntityQueryExpression()
            {
                Criteria = EntityFilter
            };
            RetrieveMetadataChangesRequest req = new RetrieveMetadataChangesRequest()
            {
                Query = entityQueryExpression,
                ClientVersionStamp = _timestamp
            };
            RetrieveMetadataChangesResponse response = (RetrieveMetadataChangesResponse)service.Execute(req);


            spinner.Stop();
            Console.WriteLine("\n Query Completed, presenting findings... \n");
            Console.ForegroundColor = ConsoleColor.White;
            StoreChangedMetaData(response.DeletedMetadata, response.EntityMetadata);




            Console.WriteLine("\n===END===");
            Console.ForegroundColor = ConsoleColor.White;

            

        }

        public static void StoreChangedMetaData(DeletedMetadataCollection deletedMD, EntityMetadataCollection changedMD)
        {
            int deletedAmount = deletedMD.Count;
            int changedAmount = changedMD.Count;

            Console.WriteLine("*Changed Metadata* \n" + deletedAmount + " deletes found. \n" + changedAmount + " changes found." );

            log("\n* CRM: "+URL+ " \n*Changed Metadata* \n" + deletedAmount + " deletes found. \n" + changedAmount + " changes found.", "Results", true, false);
            string[,] _deletedMD = new string[deletedAmount,2];
            string[,] _modifiedMD = new string[changedAmount,2];

            //Parse into Arrays
            //deletes
            int i = 0;
            foreach(var _key in deletedMD.Keys)
            {
                DataCollection<Guid> _tmpGUID;

                if (deletedMD.TryGetValue(_key, out _tmpGUID))
                {
                    Console.WriteLine(_tmpGUID);
                }else
                {
                    Console.WriteLine(_key + " not found i deleted metadata!");
                }
            }
            //changes
            for (i=0;i<changedMD.Count;i++)
            {
                Console.WriteLine("change found to: '" + changedMD[i].SchemaName.ToString() + "' since " + _ClientVersionStamp);
                log("change found to: '" + changedMD[i].SchemaName.ToString() + "' since " + _ClientVersionStamp, "Results", true, false);
            }

            
        }
        static string GetStoredTimestamp()
        {
            string _path = @"./Logs/TimeStamp.txt";
            string _ts;

            if (!File.Exists(_path))
            {
                throw new InvalidOperationException("Timestamp file does not exist, please run the 'initquery' command.");
            }

            using(StreamReader sr = new StreamReader(_path))
            {
                _ts = sr.ReadLine();
            }
            return _ts;
        }
        public static bool log(string message, string filename, bool AddTimestamp, bool OverWrite)
        {
            try
            {
                String log_item;

                DateTime timestamp = DateTime.Now;
                if (AddTimestamp)
                {
                    log_item = (timestamp + ": " + message);
                }else
                {
                    log_item = (message);
                }

                //TODO: differant file per day.
                // eg name: yyyy.mm.dd.txt

                string path = @"./Logs/" + filename + ".txt"; //ReturnedTimestamps.txt";
                if (!File.Exists(path))
                {
                    File.Create(path);
                    using (var tw = new StreamWriter(path, true))
                    {
                        tw.Write(log_item);
                        tw.Close();
                    }
                }
                else if (File.Exists(path))
                {
                    using (var tw = new StreamWriter(path, true))
                    {
                        if (!OverWrite)
                        {
                            tw.WriteLine(log_item);
                            tw.Close();

                        }else
                        {
                            tw.Write(log_item);
                            tw.Close();
                        }

                    }
                }
                //Write to log file if exists, if it doesnt exist, create one. 
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        protected static string ReadPassword()
        {
            string pass = "";
            do
            {
                ConsoleKeyInfo key = Console.ReadKey(true);
                // Backspace Should Not Work
                if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter)
                {
                    pass += key.KeyChar;
                    Console.Write("*");
                }
                else
                {
                    if (key.Key == ConsoleKey.Backspace && pass.Length > 0)
                    {
                        pass = pass.Substring(0, (pass.Length - 1));
                        Console.Write("\b \b");
                    }
                    else if (key.Key == ConsoleKey.Enter)
                    {
                        Console.WriteLine(" ");
                        return pass;
                    }
                }
            } while (true);
        }
        #endregion
    }
}
