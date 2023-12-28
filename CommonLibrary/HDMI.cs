/*
* CaptainWin.Common - Common API for test items
* HDMI.cs - Common test operations for test items
* 
* CopyRight (c) Quanta. All Rights Reserved.
*
* Authors:
*  Edison Lin  <Edison.Lin@quantatw.com>
*/

using System;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Management;
using Microsoft.Win32;

namespace CaptainWin.CommonAPI {
    public class HDMI {
        /// <summary>
        /// TitleLog
        /// </summary>
        public static void TitleLog(string content) {
           using (StreamWriter writer = new StreamWriter("c:\\TestManager\\ItemDownload\\HDMI.log", true))
           {
               writer.Write("\n[[ "+DateTime.Now.ToString()+" ]] -- "+content+" --\n");
           }
        }
        /// <summary>
        /// Log
        /// </summary>
        public static void ProcessLog(string content) {
            try {
                // appand content
                using (StreamWriter writer = new StreamWriter("c:\\TestManager\\ItemDownload\\HDMI.log", true))
                {
                    writer.Write(content+'\n');
                }
            }
            catch (Exception ex) {
                Console.WriteLine("Error!!! " + ex.Message);
            }
        }      

        public static void HdmiConnectionStatus()
        {
            TitleLog("HdmiConnectionStatus");
            int HDMI_Monitors = 0;
            ManagementClass mClass = new ManagementClass(@"\\localhost\ROOT\WMI:WmiMonitorConnectionParams");
            // ProcessLog("HdmiConnectionStatus ManagementClass ");
            
            if(mClass != null) {
                // ProcessLog(mClass);
                foreach (ManagementObject mObject in mClass.GetInstances())
                    {
                        var ss = mObject["VideoOutputTechnology"];
                        ProcessLog("HdmiConnectionStatus HDMI port Info :" + ss);
                        if (ss.ToString().StartsWith("5"))
                        {
                            int HDMIport = Convert.ToInt32(ss);
                            if (HDMIport == 5)
                            {
                                HDMI_Monitors += 1;
                            }
                        }
                    }
            }
            else
            {
                ProcessLog(" HdmiConnectionStatus Null ManagementClass ");
            }

            ProcessLog("HDMI_Monitors---> "+ HDMI_Monitors.ToString());
            // return HDMI_Monitors;
            
        }        
    }
}
