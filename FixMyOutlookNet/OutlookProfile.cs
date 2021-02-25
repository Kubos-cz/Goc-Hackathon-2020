using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;

namespace FixMyOutlookNet
{
    static class OutlookProfile
    {
        public static List<String> GetProfiles(string outlookFolderPath)
        {
            List<String> outlookProfiles = new List<string>();

            // Application returns a list of user profiles based on .ost files present in documents folder.
            // This was to be used for potential old profile and disk space cleanup
            // CURRENTLY UNUSED

            DirectoryInfo outlookFolder = new DirectoryInfo(outlookFolderPath);
            var files = outlookFolder.GetFiles("*.ost");

            foreach (FileInfo file in files)
            {
                outlookProfiles.Add(file.Name);
            }

            return outlookProfiles;
        }
        public static void ProfileCleanup(string registryPath)
        {
            Console.WriteLine($"{Localization.GetUIText(6)}");

            // Application checks if the user profile with automatically generated name for today already exists
            RegistryKey registry = Registry.CurrentUser;
            var profile = registry.OpenSubKey(registryPath, true);

            if (profile != null)
                registry.DeleteSubKeyTree(registryPath);

            registry.Close();
        }
        public static void CreateProfile(PowerShell powershellInstance, string registryPath, string profileName)
        {
            Console.WriteLine($"{Localization.GetUIText(7)}");

            // Application creates a new registry key to be used by outlook as a new profile
            powershellInstance.AddScript($"New-Item -Path {registryPath} -Name {profileName}");
            powershellInstance.Invoke();
        }
    }
}
