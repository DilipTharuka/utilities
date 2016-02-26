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
            OClient.CreateDatabasePool("localhost", 2424, "ActivityMonitor", ODatabaseType.Graph, "admin", "admin", 10, "ActivityMonitorAlias");
            oDatabase = new ODatabase("ActivityMonitorAlias");
        }

        //using(ODatabase database = new ODatabase("ModelTestDBAlias"))
        //{
        //    // prerequisites
        //    database
        //      .Create.Class("TestClass")
        //      .Extends<OVertex>()
        //      .Run();

        //    OVertex createdVertex = database
        //      .Create.Vertex("TestClass")
        //      .Set("foo", "foo string value")
        //      .Set("bar", 12345)
        //      .Run();
        //}

        public static DBConnector getInstance()
        {
            if (dbConnector == null)
                dbConnector = new DBConnector();
            return dbConnector;
        }

        public void addAppToBucket(string bucketName,string appName)
        {
            oDatabase.Query("");
        }

        public void createNewBucket(string bucketName)
        {
            oDatabase.Command("INSERT INTO Bucket(name) VALUES('" + bucketName + "')");
        }      
        
    }
}
