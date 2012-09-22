using System;
using System.Diagnostics;
using System.ServiceProcess;
using Microsoft.Win32;
using System.Timers;
using System.IO;
using System.Security;

namespace SchedulerService
{

    partial class ScheduleService : ServiceBase
    {
        public ScheduleService()
        {
            InitializeComponent();

            CanPauseAndContinue = true;

            InitializeTimer();
        }

        /// <summary>
        /// Set things in motion so your service can do its work.
        /// </summary>
        protected override void OnStart(String[] args)
        {
            // TODO: Add code here to start your service.
            base.OnStart(args);

            StartTimer();
        }

        /// <summary>
        /// Stop this service.
        /// </summary>
        protected override void OnStop()
        {
            StopTimer();
            base.OnStop();
        }

        protected override void OnPause()
        {
            base.OnPause();
        }

        protected override void OnContinue()
        {
            base.OnContinue();
        }

        protected override bool OnPowerEvent(PowerBroadcastStatus powerStatus)
        {
            switch (powerStatus)
            {
            case PowerBroadcastStatus.BatteryLow:
                break;
            case PowerBroadcastStatus.OemEvent:
                break;
            case PowerBroadcastStatus.PowerStatusChange:
                break;
            case PowerBroadcastStatus.QuerySuspend:
                break;
            case PowerBroadcastStatus.QuerySuspendFailed:
                break;
            case PowerBroadcastStatus.ResumeAutomatic:
                break;
            case PowerBroadcastStatus.ResumeCritical:
                break;
            case PowerBroadcastStatus.ResumeSuspend:
                break;
            case PowerBroadcastStatus.Suspend:
                break;
            }
            return base.OnPowerEvent(powerStatus);
        }

        protected override void OnShutdown()
        {
            CloseTimer();
            base.OnShutdown();
        }

        #region Timer

        private Timer timer;

        private void InitializeTimer()
        {
            timer = new Timer
                    {
                        Interval = 60000,
                    };

            timer.Elapsed += TimeElapsed;
        }

        private void StartTimer()
        {
            // set the timer interval and start the service
            timer.AutoReset = true;
            timer.Start();
        }

        private void StopTimer()
        {
            timer.AutoReset = false;
            timer.Stop();
        }

        private void CloseTimer()
        {
            timer.Close();
        }
        #endregion

        /// <summary>
        /// When the timer is trigered check if there are any programs to run
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TimeElapsed(Object sender, ElapsedEventArgs e)
        {
            //EventLog.WriteEntry("ScheduleService Example Timer Function called");

            try
            {
                var softwareKey = Registry.LocalMachine.OpenSubKey("Software");
                if (softwareKey == null)
                {
                    EventLog.WriteEntry("Unable to open the registry Software key for 'ScheduleService'");
                }
                else
                {
                    var key = softwareKey.OpenSubKey("ScheduleService");
                    if (key == null)
                    {
                        EventLog.WriteEntry("Unable to open the registry ScheduleExample for 'ScheduleService'");
                    }
                    else
                    {
                        try
                        {
                            var subKeyNames = key.GetSubKeyNames();

                            var dtNow = DateTime.Now;

                            foreach (var keyName in subKeyNames)
                            {
                                using (var subKey = key.OpenSubKey(keyName))
                                {
                                    if (subKey == null)
                                        continue;

                                    // get the time that the app is supposed to run and compare it to the current time.

                                    if (dtNow.Hour == Int32.Parse(subKey.GetValue("Hours").ToString())
                                     && dtNow.Minute == Int32.Parse(subKey.GetValue("Mins").ToString()))
                                    {
                                        var fileName = subKey.GetValue("FileToRun").ToString();
                                        StartProcess(fileName);
                                    }

                                    subKey.Close();
                                }
                            }
                        }
                        catch (SecurityException expSec)
                        {
                            EventLog.WriteEntry("Security exception thrown getting the sub keys names " + expSec.Message);
                        }
                        catch (IOException expIO)
                        {
                            EventLog.WriteEntry("IO exception thrown getting the sub key names " + expIO.Message);
                        }

                        key.Close();
                    }
                }
            }
            catch (ArgumentNullException expArgNull)
            {
                EventLog.WriteEntry("Argument null exception thrown " + expArgNull.Message);
            }
            catch (ArgumentException expArg)
            {
                EventLog.WriteEntry("Argument exception thrown " + expArg.Message);
            }
            catch (IOException expIO)
            {
                EventLog.WriteEntry("IO Exception thrown " + expIO.Message);
            }
            catch (SecurityException expSec)
            {
                EventLog.WriteEntry("Security exception thrown " + expSec.Message);
            }
        }

        private void StartProcess(String fileName)
        {
            var startInfo = new ProcessStartInfo(fileName) { UseShellExecute = true };

            var process = new Process { StartInfo = startInfo };

            // start the process
            // Notice that here I am not keeping track of the processes just letting them run.
            if (!process.Start())
            {
                EventLog.WriteEntry("Unable to start the process " + startInfo.FileName);
            }
        }
    }
}
