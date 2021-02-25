using System;
using System.Management.Automation;


namespace FixMyOutlookNet
{
    class Application
    {
        // Application main function
        static void Main(string[] args)
        {

            //Inform user
            Console.WriteLine("*******************************");
            Console.WriteLine($"{Localization.GetUIText(1)}");
            Console.WriteLine("*******************************");
            Console.WriteLine("");

            try
            {
                // Close outlook warning
                Console.WriteLine($"{Localization.GetUIText(2)}");
                Console.ReadKey();

                // Create powershell instance
                using (PowerShell powerShell = PowerShell.Create(RunspaceMode.NewRunspace))
                {
                    // Close Outlook
                    OutlookApplication.CloseOutlook(powerShell);

                    // Get my office version
                    string officeVersion = OutlookApplication.GetMyOfficeVersion(powerShell);

                    // Set registry path
                    Console.WriteLine($@"{Localization.GetUIText(5)} {officeVersion}");
                    string registryPath = $@"software\microsoft\office\{officeVersion}.0\outlook";
                    
                    // Set the new profile name
                    string profileName = $"{System.Security.Principal.WindowsIdentity.GetCurrent().Name.Replace("GROUPHC\\", "")}_{DateTime.Now.ToString("dd-MM-yyyy")}";

                    // Cleanup old auto-generated profile
                    OutlookProfile.ProfileCleanup($@"{registryPath}\profiles\{profileName}");

                    //Create new profile
                    OutlookProfile.CreateProfile(powerShell, $@"HKCU:\{registryPath}\profiles", profileName);

                    // Restart outlook with new profile
                    OutlookApplication.SetDefaultProfile($@"HKEY_CURRENT_USER\{registryPath}", profileName);
                }

                // Inform user
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"{Localization.GetUIText(9)}");
                Console.ReadKey();
            }
            catch(Exception x)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Outlook configuration has failed, press any key to continue");
                Console.ReadKey();
            }
        }

    }
}
