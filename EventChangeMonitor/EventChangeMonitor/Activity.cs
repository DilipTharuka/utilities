using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace ActivityMonitor
{
    [Serializable()]
    class Activity
    {
        public DateTime startTime { get; set; } 
        public DateTime endTime { get; set; }
        public string processName { get; set; }
        public TimeSpan duration { get; set; }

        public void generateDuration()
        {
            this.duration += endTime.Subtract(startTime);
        }

    }
}
