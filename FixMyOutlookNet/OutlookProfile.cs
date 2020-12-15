using System;
using System.IO;
using Microsoft.Win32;
using System.Collections.Generic;
using System.Management.Automation;


namespace FixMyOutlookNet
{
    class OutlookProfile
    {
        // Application main function
        static void Main(string[] args)
        {
            //Inform user
            Console.WriteLine("*******************************");
            Console.WriteLine("|| OUTLOOK CONFIGURATION FIX ||");
            Console.WriteLine("*******************************");
            Console.WriteLine("");

            try
            {
                // Close outlook warning
                Console.WriteLine("Outlook will be closed to proceed, press any key to continue");
                Console.ReadKey();

                // Create powershell instance
                using (PowerShell powerShell = PowerShell.Create(RunspaceMode.NewRunspace))
                {
                    // Close Outlook
                    CloseOutlook(powerShell);

                    // Get my office version
                    string officeVersion = GetMyOfficeVersion(powerShell);

                    // Set registry path
                    Console.WriteLine($@"Detected version: {officeVersion}");
                    string registryPath = $@"software\microsoft\office\{officeVersion}.0\outlook";
                    
                    // Set the new profile name
                    string profileName = $"{System.Security.Principal.WindowsIdentity.GetCurrent().Name.Replace("GROUPHC\\", "")}_{DateTime.Now.ToString("dd-MM-yyyy")}";

                    // Cleanup old auto-generated profile
                    ProfileCleanup($@"{registryPath}\profiles\{profileName}");

                    //Create new profile
                    CreateProfile(powerShell, $@"HKCU:\{registryPath}\profiles", profileName);

                    // Restart outlook with new profile
                    SetDefaultProfile($@"HKEY_CURRENT_USER\{registryPath}", profileName);
                }

                // Inform user
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Outlook configuration has finished you may now start outlook, press any key to continue");
                Console.ReadKey();
            }
            catch(Exception x)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Outlook configuration has failed, press any key to continue");
                Console.ReadKey();
            }
        }

        // Helper functions
        static List<String> GetProfiles(string outlookFolderPath)
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
        static void CloseOutlook(PowerShell powerShellInstance)
        {
            Console.WriteLine("Closing outlook...");

            // Application closes the outlook process
            powerShellInstance.AddScript(@"Stop-Process -Name OUTLOOK");
            powerShellInstance.Invoke();
        }
        static string GetMyOfficeVersion(PowerShell powerShellInstance)
        {
            Console.WriteLine("Detecting office instalation...");

            // Very basic outlook version detection. 
            // Application assumes that the default version of office is '15' and if an office package referencing office 365 is detected the application assumes the user has the '16' version
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
        static void ProfileCleanup(string registryPath)
        {
            Console.WriteLine("Cleaning auto-generated profiles...");

            // Application checks if the user profile with automatically generated name for today already exists
            RegistryKey registry = Registry.CurrentUser;
            var profile = registry.OpenSubKey(registryPath, true);

            if (profile != null)
                registry.DeleteSubKeyTree(registryPath);

            registry.Close();
        }
        static void CreateProfile(PowerShell powershellInstance,string registryPath, string profileName)
        {
            Console.WriteLine("Creating profile...");

            // Application creates a new registry key to be used by outlook as a new profile
            powershellInstance.AddScript($"New-Item -Path {registryPath} -Name {profileName}");
            powershellInstance.Invoke();
        }
        static void SetDefaultProfile(string registryPath, string profileName)
        {
            Console.WriteLine("Setting up outlook configuration...");

            // Application overwrites outlook settings to ensure the new profile is loaded on outlook start
            Registry.SetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Exchange\Client\Options", "PickLogonProfile", "0", RegistryValueKind.String);
            Registry.SetValue(registryPath, "DefaultProfile", profileName, RegistryValueKind.String);
        }
    }
}
