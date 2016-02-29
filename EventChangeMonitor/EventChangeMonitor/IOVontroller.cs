using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActivityMonitor
{
    class IOController
    {
        public void Serialize(Dictionary<string, Activity> dictionary)
        {
            var f_fileStream = new FileStream(DateTime.Now.ToString("yyyy-MM-dd") + ".amos", FileMode.Create, FileAccess.Write);
            var f_binaryFormatter = new System.Runtime.Serialization.Formatters.Binary.BinaryFormatter();
            f_binaryFormatter.Serialize(f_fileStream, dictionary);
            f_fileStream.Close();
        }

        public Dictionary<string, Activity> Deserialize()
        {
            var f_fileStream = File.OpenRead(@"dictionarySerialized.xml");
            var f_binaryFormatter = new System.Runtime.Serialization.Formatters.Binary.BinaryFormatter();
            Dictionary<string, Activity> dictionary = (Dictionary<string, Activity>)f_binaryFormatter.Deserialize(f_fileStream);
            f_fileStream.Close();
            return dictionary;
        }
    }
}
