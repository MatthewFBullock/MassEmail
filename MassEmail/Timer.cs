using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace MassEmail
{
    class Timer
    {
        private static System.Timers.Timer aTimer;

        public static void ReferenceTimer(ref DateTime _aTimer)
        {
            SetTimer();

            
            Console.WriteLine("The application started at {0:HH:mm:ss}", DateTime.Now);
            aTimer.Stop();
            aTimer.Dispose();
        }

        private static void SetTimer()
        {
            // Create a timer with a two second interval.
            aTimer = new System.Timers.Timer(15000);
            // Hook up the Elapsed event for the timer. 
            aTimer.Elapsed += OnTimedEvent;
            aTimer.AutoReset = true;
            aTimer.Enabled = true;



        }

        public static void OnTimedEvent(object sender, ElapsedEventArgs e)
        {
            Console.WriteLine("The Elapsed event was raised at {0:HH:mm:ss}", e.SignalTime);
        }
    }
}
