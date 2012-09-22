using System;
using System.IO;
using System.Security;
using Microsoft.Win32;

namespace TestInstallerCode
{
    /// <summary>
    /// Summary description for Class1.
    /// </summary>
    class InstallerTest
    {
        public void Install()
        {
            try
            {
                // just create the key the gui part of the code will update it
                var reg = Registry.CurrentUser.OpenSubKey("Software", true);
                if (null == reg)
                {
                    Console.WriteLine("Error trying to install 'ScheduleExample'");
                }
                else
                {
                    var scheduleKey = reg.CreateSubKey("ScheduleExample");
                    if (null == scheduleKey)
                    {
                        Console.WriteLine("Error trying to install 'ScheduleExample'");
                    }
                    else
                    {
                        scheduleKey.Close();
                    }
                    reg.Close();
                }
            }
            catch (ArgumentNullException argNullExp)
            {
                Console.WriteLine("Error with the argument subkey " + argNullExp.Message);
            }
            catch (SecurityException secExp)
            {
                Console.WriteLine("Error the user does not have access permission " + secExp.Message);
            }
            catch (IOException ioExp)
            {
                Console.WriteLine("Error the registry key is closed " + ioExp.Message);
            }
            catch (UnauthorizedAccessException unExp)
            {
                Console.WriteLine("Error the user does not have access permission " + unExp.Message);
            }
            catch (ArgumentException argExp)
            {
                Console.WriteLine("Error in the install data format " + argExp.Message);
            }
            catch (Exception exp)
            {
                Console.WriteLine("A problem occured with the install " + exp.Message);
            }
        }

        /// <summary>
        /// override the uninstall method and remove the registry key
        /// </summary>
        /// <param name="iInstallData"></param>
        public void Uninstall()
        {
            try
            {
                var openSubKey = Registry.CurrentUser.OpenSubKey("Software", true);
                if (null != openSubKey) openSubKey.DeleteSubKeyTree("ScheduleExample");

                Console.WriteLine("'ScheduleExample' has been uninstalled");
            }
            catch (ArgumentException argExp)
            {
                Console.WriteLine("Error in the install data format " + argExp.Message);
            }
        }

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //
            // TODO: Add code to start application here
            //
            var installerTest = new InstallerTest();

            installerTest.Install();

            installerTest.Uninstall();
        }
    }
}