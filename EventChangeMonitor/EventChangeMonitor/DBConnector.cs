using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Orient.Client;

namespace ActivityMonitor
{
    class DBConnector
    {
        private static DBConnector dbConnector = null;
        private static ODatabase oDatabase = null;

        private DBConnector()
        {
            OClient.CreateDatabasePool("cmddtkarunathil", 2424, "ActivityMonitor", ODatabaseType.Graph, "admin", "admin", 10, "ActivityMonitorAlias");
            oDatabase = new ODatabase("ActivityMonitorAlias");
        }

        public static DBConnector getInstance()
        {
            if (dbConnector == null)
                dbConnector = new DBConnector();
            return dbConnector;
        }

        public void addAppToBucket(string bucketName, string appName)
        {
            Console.WriteLine("UPDATE Bucket ADD applications='" + appName + "' WHERE name = '" + bucketName + "'");
            oDatabase.Command("UPDATE Bucket ADD applications='"+appName+"' WHERE name = '"+bucketName+"'");    
        }

        public void createNewBucket(string bucketName)
        {
            oDatabase.Command("INSERT INTO Bucket(name) VALUES('" + bucketName + "')");
        }

        public Dictionary<string, List<string>> getBuckets()
        {
            Dictionary<string, List<string>> buckets = new Dictionary<string, List<string>>();
            List<ODocument> oDocuments =  oDatabase.Query("SELECT FROM Bucket");
            foreach (ODocument oDocument in oDocuments)
            {
                buckets.Add(oDocument.GetField<String>("name"), oDocument.GetField<List<String>>("applications"));
            }

            //foreach (KeyValuePair<string, List<string>> pair in buckets)
            //{
            //    Console.WriteLine(pair.Key);
            //    foreach (string item in pair.Value)
            //    {
            //        Console.Write(item + " ");
            //    }
            //}

            return buckets;
        }

        public List<string> getPackages()
        {
            List<string> packages = new List<string>();
            List<ODocument> oDocuments = oDatabase.Query("SELECT FROM Package");
            foreach (ODocument oDocument in oDocuments)
            {
                packages.Add(oDocument.GetField<String>("name"));
            }
            return packages;

        }

    }
}
