using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyCommonUtilities
{
    public class MySimpleDurationTimer
    {
        TimeSpan _stop;
        TimeSpan _start;

        public MySimpleDurationTimer()
        {
            _start = new TimeSpan(DateTime.Now.Ticks);
        }

        public TimeSpan StopWithDuration()
        {
            _stop = new TimeSpan(DateTime.Now.Ticks);
            return _stop.Subtract(_start);
        }
    }
}
