/*
* CaptainWin.Common - Common API for test items
* GetDark.cs - Common test operations for test items
* 
* CopyRight (c) Quanta. All Rights Reserved.
*
* Authors:
*  Edison Lin  <Edison.Lin@quantatw.com>
*/

using System;
using System.IO;
using Microsoft.Win32;

namespace CaptainWin.CommonAPI {
    /// <summary>
    /// GetDark
    /// </summary>
    public class GetDark {
        /// <summary>
        /// TitleLog
        /// </summary>
        public static void TitleLog(string content) {
           using (StreamWriter writer = new StreamWriter("c:\\TestManager\\ItemDownload\\GetDark.log", true))
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
                using (StreamWriter writer = new StreamWriter("c:\\TestManager\\ItemDownload\\GetDark.log", true))
                {
                    writer.Write(content+'\n');
                }

            }
            catch (Exception ex) {
                Console.WriteLine("Error!!! " + ex.Message);
            }
        }      
        /// <summary>
        /// GetDark UI
        /// </summary>
        public static void IsDark() {
            TitleLog("IsDark");
            const string keyName = @"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";
            const string valueName = "AppsUseLightTheme";

            try
            {
                // Read registry
                object registryValue = Registry.GetValue(keyName, valueName, null);

                if (registryValue != null)
                {
                    // 0 = Dark
                    bool isDarkMode = (int)registryValue == 0;
                    ProcessLog($"Is Dark Mode Enabled: {isDarkMode}");
                }
                else
                {
                    ProcessLog("Unable to retrieve the registry value.");
                }
            }
            catch (Exception ex)
            {
                ProcessLog($"Error: {ex.Message}");
            }
        }
    }
}
