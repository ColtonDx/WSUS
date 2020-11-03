<#
 __      _____ _   _ ___ 
 \ \    / / __| | | / __|
  \ \/\/ /\__ \ |_| \__ \
   \_/\_/ |___/\___/|___/
                                                       



Description:
    This script does the following:
        1. Pulls the domain information from the PC its run from.
        2. Declines updates from a CSV file that you upload to C:\Admin\WSUS
        3. Declines unwanted updates like language packs
        4. Approves updates older than 6 months.
#>

<#
 __   ___   ___ ___   _   ___ _    ___ ___ 
 \ \ / /_\ | _ \_ _| /_\ | _ ) |  | __/ __|
  \ V / _ \|   /| | / _ \| _ \ |__| _|\__ \
   \_/_/ \_\_|_\___/_/ \_\___/____|___|___/
                                           
This section starts all the variables and objects that will be used when the script runs. It also makes the 
log file.
#>
 
 #Pull Computer Name from Current PC (Script should be ran fmor WSUS Host)
$Computer = $env:COMPUTERNAME

 #Pull Domain Name from Current PC (must be on same Domain as WSUS)
$Domain = $env:USERDNSDOMAIN

 #Generate Fully Qualified Domain Name of Current PC
$FQDN = "$Computer" + "." + "$Domain"

 #Set the update server to the FQDN
[String]$updateServer1 = $FQDN

 #Disable Secure Connection, enable if this is set in GPO
[Boolean]$useSecureConnection = $False

 #Only change if changed in GPO
[Int32]$portNumber = 8530

 #Load CSV for Wonderware. KBs in this CSV will be declined.
$path = "C:\Admin\WSUS\kb.csv"

#Create Variable for KB#s
$WWDecline = cat $path

 #Generate a Log File
$log = "C:\Admin\WSUS\Approved_Updates_{0:MMddyyyy_HHmm}.log" -f (Get-Date)
new-item -path $log -type file -force

#Load Prerequisites 
[void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")

 #Declined-Count
 $dcount = 0
 
 #Update-Count
 $ucount = 0

 # Connect to WSUS Server
 $updateServer = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($updateServer1,$useSecureConnection,$portNumber)
  $updatescope = New-Object Microsoft.UpdateServices.Administration.UpdateScope

 #Update Scope to Only be "Not Approved"
 $u=$updateServer.GetUpdates($updatescope ) | ? {($_.IsDeclined -eq $False)}
 
 



 <#
  _    ___   ___  ___  ___ 
 | |  / _ \ / _ \| _ \/ __|
 | |_| (_) | (_) |  _/\__ \
 |____\___/ \___/|_|  |___/
This section runs the loops, this is the actual meat and potatoes of the script. Each Loop Will Approve or Deny Updates on parameters.
#>
 # Connect to WSUS Server
 $updateServer = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($updateServer1,$useSecureConnection,$portNumber)
 write-host "<<<Connected successfully >>>" -foregroundcolor "yellow"
 $updatescope = New-Object Microsoft.UpdateServices.Administration.UpdateScope

 #Update Scope to Only be "Not Approved"
 $u=$updateServer.GetUpdates($updatescope ) | ? {($_.IsDeclined -eq $False)}
 #This declares the scope, updates that are not declined and not approved
#########################
# Decline by Title      #
#########################
#Decline Updates that have titles matching undesired updates

#For Loop to Decline 
write-host Declining Updates based on Title
foreach ($u1 in $u )
 { 
  if ($u1.Title -Like '*x86*'  -or $u1.Title -Like '*SkyDrive*' -or $u1.Title -Like '*Japan*' -or $u1.Title -Like '*Korean*' -or $u1.Title -Like '*OneNote*' -or $u1.Title -Like '*ARM64*' -or $u1.Title -Like '*OneDrive*'  -or $u1.Title -Like '*Office 2010*' -or $u1.Title -Like '*LanguagePack*' -or $u1.Title -Like '*Insider Preview*' -or $u1.Title -Like '*FeatureOnDemand*' -or $u1.Title -Like '*Feature On Demand*' -or $u1.Title -Like '*Lang Pack*' -or $u1.Title -Like '*Excel Web App*'  -or $u1.Title -Like '*Windows Server Next*'  -or $u1.Title -Like '*Windows 10 Version Next*' -or $u1.Title -Like '*Sharepoint Workspace*' -or $u1.Title -Like '*farm-deployment*' -or $u1.Title -Like '*Visio 2010 Viewer*')
  #If the update is: LanguagePacks, x86, Insider Previews, FeatureOnDemand, and ARM then proceed
    {
      write-host Decline Update : $u1.Title
      $u1.Decline()
      $dcount=$dcount + 1
    }
 }
  # Connect to WSUS Server
 $updateServer = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($updateServer1,$useSecureConnection,$portNumber)
  $updatescope = New-Object Microsoft.UpdateServices.Administration.UpdateScope

 #Update Scope to Only be "Not Approved"
 $u=$updateServer.GetUpdates($updatescope ) | ? {($_.IsDeclined -eq $False)}
 
#########################
# Decline               #
#########################
#Declines Updates located in C:\Admin\KB.CSV

 write-host Declining Updates from KB.csv
foreach ($u1 in $u)
{
 foreach ($WWDecline1 in $WWDecline )
  {
  $WWU = $WWDecline1.Insert(0,"*")
  $WWU = $WWU+"*"
   if ($u1.Title -Like $WWU)
    {
      $u1.Decline()
      write-host Declined: $u1.Title
      $dcount=$dcount + 1
  }
}
}
  # Connect to WSUS Server
 $updateServer = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($updateServer1,$useSecureConnection,$portNumber)

 $updatescope = New-Object Microsoft.UpdateServices.Administration.UpdateScope

 #Update Scope to Only be "Not Approved"
 $u=$updateServer.GetUpdates($updatescope ) | ? {($_.IsDeclined -eq $False)}

#########################
# Approve >6mo.         #
#########################
#Approve Updates Older than Six months

   $wgroups = $updateserver.GetComputerTargetGroups()  | ? {$_.Name -eq "RobinsonLake"}
   foreach ($wgroup in $wgroups)
   {
   foreach ($u1 in $u)
     {
	  #Only look for Updates older than 6 months
	  if ($u1.CreationDate -le ((get-date).AddMonths(-6)))
	    {
	     #Approve Updates for Current Group
         $u1.Approve("Install", $wgroup)
	     write-host Approved Update : $u1.Title
         #write-host $u1.CreationDate
         #write-host ((get-Date).AddMonths(-6))
	     $ucount=$ucount + 1
	    }
     }
   }
 



<#
  _    ___   ___ ___ 
 | |  / _ \ / __/ __|
 | |_| (_) | (_ \__ \
 |____\___/ \___|___/
 This section writes the log file
#>

 #Declines Done
write-host Total Declined Updates: $dcount
 #Updates Approved
write-host Total Approved Updates: $ucount
 #Planning on adding a log here to keep track of what is approved/declined and when
 
 #Get Todays Date
 $date = Get-Date
 "Aproved updates (on " + $date + "): " | Out-File $log -append
 #List Approved Updates
"Updates have been approved for following groups: (" + $groups + ")" | Out-File $log -append
 #Add Updates to Log File
"Folowing updates have been approved:" | Out-File $log -append 
$u | Select Title,ProductTitles,KnowledgebaseArticles,CreationDate | ft -Wrap | Out-File $log -append
 
 #Errors
{
 
write-host "Error Occurred"
 write-host "Exception Message: "
 write-host $_.Exception.Message
 write-host $_.Exception.StackTrace
 exit
 }
 