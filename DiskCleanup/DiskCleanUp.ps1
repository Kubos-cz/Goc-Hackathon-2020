##################################################################################   
<# 
 
This script is created to automate the cleanup activity. Doing so will benefit to reduce the size of disk.
 
This script will perform the following 
1. Clear windows temp and user temp folder 
2. Empty recycle bin 
3. Disk Cleanup 
4. Clear CBS cabinet log files 
5. Clear downloaded patches 
6. Clear downloaded driver 
7. Clean download folder 

 
Note: 
 
1. Run the script with Administrative access. 
2. Put # chanracter before delete() function if you want to skip any folder. 
3. Diskcleanup utility will prompt to select options available. Select the required options before 30 sec. 
4. To add any new folder , need to declare the folder and call delete function with that folder variable 
 
contact: SHISHIR KUSHAWAHA (srktcet@gmail.com) 
 
#> 
##################################################################################   
  
## Function ## 
 
        #1# This function will convert byte Data to megabyte. 
 
         function foldersize($folder) 
         { 
 
         $folderSizeinbyte = (Get-ChildItem $folder -Recurse | Measure-Object -property length -sum) 
   
         $folderSizeinMB=($folderSizeinbyte.sum / 1048576) 
  
         return $folderSizeinMB 
         } 
  
         #2# This function will display the folder size before deletion. 
 
         function before($folder1) 
        { 
 
            $x=foldersize($folder1) 
            write-host "Total size before deletion=$x MB" 
            return $x 
        } 
 
        #3# This function will display the folder size after deletion. 
         function post($folder2) 
        { 
 
            $y=foldersize($folder2) 
            write-host "Total size after deletion $y MB" 
            return $y 
        } 
 
        #4# This function will display the warning message. 
        function msg($folder3) 
        { 
        write-Host "Removing Junk files in $folder3." -ForegroundColor Yellow -background black 
        } 
 
 
        #5# This function will display the total spcae cleared. 
        function totalmsg($folder4,$sum) 
        { 
        write-Host "Total space cleared in MB from $folder4" $Sum  -ForegroundColor Green 
        } 
 
        ## This function will cleanup the specified folder 
        function delete($folder5) 
        { 
 
            #[double]$a=before($folder5) 
            #msg($folder5) 
            Remove-Item -Recurse  $folder5 -Force -Verbose  
            #[double]$b=post($folder5)  
 
            #$total=$a-$b 
            #totalmsg($folder5,$total) 
            #$a=0 
            #$b=0 
            #$total=0 
        } 
##End of Functions Declartion.## 
 
 
## Variables Declaration####    
    
        $objShell = New-Object -ComObject Shell.Application    
        $Recyclebin = $objShell.Namespace(0xA)      
        $temp = get-ChildItem "env:\TEMP"    
        $temp2 = $temp.Value    
        $WinTemp = "$env:SystemDrive\Windows\Temp\*"      
        $CBS="$env:SystemDrive\Windows\Logs\CBS\"  
        $swtools="$env:SystemDrive\swtools\*" 
        $drivers="$env:SystemDrive\drivers\*"
        $hp="$env:SystemDrive\hp\*"
        $dell="$env:SystemDrive\dell\*"
        $intel="$env:SystemDrive\intel\*"
        $swsetup="$env:SystemDrive\swsetup\*" 
        $downloads="$env:SystemDrive\users\ADMINI~1\downloads\*" 
        $Prefetch="$env:SystemDrive\Windows\Prefetch\*" 
        $DowloadeUpdate="$env:SystemDrive\Windows\SoftwareDistribution\Download\*" 
        $Outlook="$env:LOCALAPPDATA\Microsoft\Outlook\*" 
##End of variable Declartion.## 
    
##Execution## 

##It writes before deletion disk properties## 
Get-CimInstance -Class CIM_LogicalDisk | Select-Object @{Name="Size(GB)";Expression={$_.size/1gb}}, @{Name="Free Space(GB)";Expression={$_.freespace/1gb}}, @{Name="Free (%)";Expression={"{0,6:P0}" -f(($_.freespace/1gb) / ($_.size/1gb))}}, DeviceID, DriveType | Where-Object DriveType -EQ '3'
    
     # Remove temp files located in "C:\Users\USERNAME\AppData\Local\Temp"    
        [double]$a=before($temp2) 
        msg($temp2) 
        Remove-Item -Recurse  "$temp2\*" -Force -Verbose  
        [double]$b=post($temp2)  
 
        $total=$a-$b 
        totalmsg($temp2,$total) 
 
     
    # Remove content of folder created during installation of driver     
        delete($swtools) 
     
 
    # Remove content of folder created during installation of Lenovo driver     
        delete($drivers) 

    # Remove content of folder created during installation of Intel driver     
        delete($intel)  

    # Remove content of folder created during installation of HP driver     
        delete($hp) 
        
    # Remove content of folder created during installation of Dell driver     
       delete($dell)       
 
    # Remove content of folder created during installation of HP driver     
        delete($swsetup) 
 
    # Remove content of download folder of administrator account     
        delete($downloads)    
 
    # Empty Recycle Bin   
            write-Host "Emptying Recycle Bin." -ForegroundColor Cyan     
        $Recyclebin.items() | %{ remove-item $_.path -Recurse -verbose -Confirm:$false}    
 
 
    # Remove Windows Temp Directory  
        delete($WinTemp) 
    
 
 
    # Remove Prefetch folder content 
        delete($Prefetch) 
    
     
    # Remove CBS log file 
        delete($CBS) 
     
 
    # Remove downloaded update 
        delete($DowloadeUpdate) 


    # Remove outlook profiles
     $ost = get-ChildItem $Outlook | where { $_.Extension -eq ".ost"} and {($_.LastWriteTime -gt (Get-Date).AddDays(-14))} 
     $ost | remove-Item  
     
     # Empty recycle bin

     Clear-RecycleBin -Force

     ##It writes after deletion disk status## 
Get-CimInstance -Class CIM_LogicalDisk | Select-Object @{Name="Size(GB)";Expression={$_.size/1gb}}, @{Name="Free Space(GB)";Expression={$_.freespace/1gb}}, @{Name="Free (%)";Expression={"{0,6:P0}" -f(($_.freespace/1gb) / ($_.size/1gb))}}, DeviceID, DriveType | Where-Object DriveType -EQ '3'
 
 
##End of execution## 
##### End of the Script #####  
