using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace MyCommonUtilities
{
    class InvokeTimer
    {
        /// <summary>
        /// May Be Issues If Invoke Is Called With Another Callback Before Timer is Elapsed
        /// </summary>
        static Timer onceInvokeTimer = null;

        static Dictionary<Action, Timer> ActionTimerDictionary = new Dictionary<Action, Timer>();

        public static void Invoke(float delay, Action callback)
        {
            if (onceInvokeTimer != null)
            {
                onceInvokeTimer.Stop();
                onceInvokeTimer.Dispose();
            }

            onceInvokeTimer = new System.Timers.Timer();
            onceInvokeTimer.Interval = delay;
            onceInvokeTimer.AutoReset = false;
            onceInvokeTimer.Elapsed += (obj, elapsedArgs) =>
            {
                onceInvokeTimer.Stop();
                callback();
                onceInvokeTimer.Dispose();
            };
            onceInvokeTimer.Start();
        }

        public static void InvokeRepeating(float interval, Action callback, TaskScheduler _scheduler = null)
        {
            if (ActionTimerDictionary.ContainsKey(callback))
            {
                var _oldtimer = ActionTimerDictionary[callback];
                _oldtimer.Stop();
                _oldtimer.Dispose();
                ActionTimerDictionary.Remove(callback);
            }
            Timer newTimer = new Timer();
            newTimer.Interval = interval;
            newTimer.AutoReset = true;
            newTimer.Elapsed += (obj, elapsedArgs) =>
            {
                if (_scheduler != null)
                {
                    Task _task = new Task(callback);
                    _task.Start(_scheduler);
                }
                else
                {
                    callback();
                }
            };
            ActionTimerDictionary.Add(callback, newTimer);
            newTimer.Start();
        }

        public static void CancelInvoke()
        {
            foreach (Timer timer in ActionTimerDictionary.Values)
            {
                timer.Stop();
                timer.Dispose();
            }
            ActionTimerDictionary = new Dictionary<Action, Timer>();
        }
    }
}
