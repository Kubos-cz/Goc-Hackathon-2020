using Microsoft.Win32;
using System;
using System.Linq;
using System.Management.Automation;

namespace FixMyOutlookNet
{
    class OutlookApplication
    {
        public static void CloseOutlook(PowerShell powerShellInstance)
        {
            Console.WriteLine($"{Localization.GetUIText(3)}");

            // Application closes the outlook process
            powerShellInstance.AddScript(@"Stop-Process -Name OUTLOOK");
            powerShellInstance.Invoke();
        }
        public static string GetMyOfficeVersion(PowerShell powerShellInstance)
        {
            Console.WriteLine($"{Localization.GetUIText(4)}");

            // Very basic outlook version detection. 
            // Application assumes that the default version of office is '15' and if an office package referencing office 365 is detected the application assumes the user has the '16' version
            // Win32Reg_AddRemovePrograms is not a standard Windows class. This WMI class is only loaded during the installation of an SMS/SCCM client, use Win32_Product instead

            powerShellInstance.AddScript(@"Get-WmiObject Win32Reg_AddRemovePrograms | where{$_.DisplayName -like ""Office 365*""} | select DisplayName,Version");
            var result = powerShellInstance.Invoke();

            if (result.Count > 0)
            {
                return "16";
            }
            else
            {
                return "15";
            }
        }
        public static void SetDefaultProfile(string registryPath, string profileName)
        {
            Console.WriteLine($"{Localization.GetUIText(8)}");

            // Application overwrites outlook settings to ensure the new profile is loaded on outlook start
            Registry.SetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Exchange\Client\Options", "PickLogonProfile", "0", RegistryValueKind.String);
            Registry.SetValue(registryPath, "DefaultProfile", profileName, RegistryValueKind.String);
        }
    }
}
