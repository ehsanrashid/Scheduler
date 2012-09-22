using System;
using System.Collections;
using System.ComponentModel;
using System.Configuration.Install;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.Win32;
using System.IO;
using System.Security;

namespace SchedulerService
{
    [RunInstaller(true)]
    public partial class ScheduleInstaller : Installer
    {
        readonly EventLog _EventLog;

        public ScheduleInstaller()
        {
            InitializeComponent();

            _EventLog = new EventLog();
        }

        /// <summary>
        /// override the install method to set up the information.
        /// all thats create here is a registry key. It should be noted that this function
        /// can't be debugged so catch all possible exceptions
        /// </summary>
        /// <param name="iInstallData"></param>
        public override void Install(IDictionary iInstallData)
        {
            try
            {
                // must call base class install first
                base.Install(iInstallData);

                // just create the key the gui part of the code will update it
                using (var regSoft = Registry.LocalMachine.OpenSubKey("Software", true))
                {
                    if (null == regSoft)
                    {
                        _EventLog.WriteEntry("Error trying to install 'ScheduleService'");
                        return;
                    }

                    using (var regSchedule = regSoft.CreateSubKey("ScheduleService"))
                    {
                        if (null == regSchedule)
                        {
                            _EventLog.WriteEntry("Error trying to install 'ScheduleService'");
                            return;
                        }

                        regSchedule.Close();
                    }
                    regSoft.Close();
                }
            }
            catch (ArgumentNullException expArgNull)
            {
                _EventLog.WriteEntry("Error with the argument subkey " + expArgNull.Message);
            }
            catch (SecurityException expSec)
            {
                _EventLog.WriteEntry("Error the user does not have access permission " + expSec.Message);
            }
            catch (IOException expIO)
            {
                _EventLog.WriteEntry("Error the registry key is closed " + expIO.Message);
            }
            catch (UnauthorizedAccessException expUA)
            {
                _EventLog.WriteEntry("Error the user does not have access permission " + expUA.Message);
            }
            catch (ArgumentException expArg)
            {
                _EventLog.WriteEntry("Error in the install data format " + expArg.Message);
            }
            catch (Exception exp)
            {
                _EventLog.WriteEntry("A problem occured with the install " + exp.Message);
            }
        }

        /// <summary>
        /// override the uninstall method and remove the registry key
        /// </summary>
        /// <param name="iInstallData"></param>
        public override void Uninstall(IDictionary iInstallData)
        {
            try
            {
                base.Uninstall(iInstallData);

                //base.Uninstall(iInstallData);

                var regSoft = Registry.LocalMachine.OpenSubKey("Software", true);
                if (null != regSoft) regSoft.DeleteSubKeyTree("ScheduleService");
            }
            catch (ArgumentException expArg)
            {
                MessageBox.Show("Error in the install data format " + expArg.Message);
            }
            catch (InstallException expInst)
            {
                MessageBox.Show("A problem occurred with the install " + expInst.Message);
            }
        }
    }
}
