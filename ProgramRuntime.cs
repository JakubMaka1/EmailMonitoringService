using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailMonitoringService
{
    public class ProgramRuntime : Service1
    {
        public Stopwatch stopwatch;

        public ProgramRuntime()
        {
            stopwatch = new Stopwatch();
            stopwatch.Start();
        }

        public void StopAndDisplayRuntime()
        {
            stopwatch.Stop();
            TimeSpan runtime = stopwatch.Elapsed;
            Logger.WriteSystemLog($"Program działał przez: {runtime.Days} dni {runtime.Hours} godz {runtime.Minutes} min {runtime.Seconds} sek {runtime.Milliseconds} ms");
            SMTP.SendEmail("jakub.maka@elis.com", "Program przestał działać", $"Program działał przez: {runtime.Days} dni {runtime.Hours} godz {runtime.Minutes} min {runtime.Seconds} sek {runtime.Milliseconds} ms");
        }

        public void ErrorStopAndDisplayRuntime(string error)
        {
            stopwatch.Stop();
            TimeSpan runtime = stopwatch.Elapsed;
            Logger.WriteSystemLog($"Program działał przez: {runtime.Days} dni {runtime.Hours} godz {runtime.Minutes} min {runtime.Seconds} sek {runtime.Milliseconds} ms");
            SMTP.SendEmail("jakub.maka@elis.com", $"{error}",$"Program działał przez: {runtime.Days} dni {runtime.Hours} godz {runtime.Minutes} min {runtime.Seconds} sek {runtime.Milliseconds} ms");

        }
    }
}
