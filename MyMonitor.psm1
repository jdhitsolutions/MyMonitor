#requires -version 4.0



<#
TODO: WRITE TO LOCAL SQL DATABASE

Learn more about PowerShell:
http://jdhitsolutions.com/blog/essential-powershell-resources/

 
  ****************************************************************
  * DO NOT USE IN A PRODUCTION ENVIRONMENT UNTIL YOU HAVE TESTED *
  * THOROUGHLY IN A LAB ENVIRONMENT. USE AT YOUR OWN RISK.  IF   *
  * YOU DO NOT UNDERSTAND WHAT THIS SCRIPT DOES OR HOW IT WORKS, *
  * DO NOT USE IT OUTSIDE OF A SECURE, TEST SETTING.             *
  ****************************************************************
#>

#region Commands

Function Get-ForegroundWindowProcess {

<#
.Synopsis
Get process for foreground window process
.Description
This command will retrieve the process for the active foreground window, ignoring any process with a main window handle of 0.

It will also ignore Task Switching done with Explorer.
.Example
PS C:\> get-foregroundwindowprocess

Handles  NPM(K)    PM(K)      WS(K) VM(M)   CPU(s)     Id ProcessName                                
-------  ------    -----      ----- -----   ------     -- -----------                                
    538      57   124392     151484   885    34.22   4160 powershell_ise

.Notes

Learn more about PowerShell:
http://jdhitsolutions.com/blog/essential-powershell-resources/


  ****************************************************************
  * DO NOT USE IN A PRODUCTION ENVIRONMENT UNTIL YOU HAVE TESTED *
  * THOROUGHLY IN A LAB ENVIRONMENT. USE AT YOUR OWN RISK.  IF   *
  * YOU DO NOT UNDERSTAND WHAT THIS SCRIPT DOES OR HOW IT WORKS, *
  * DO NOT USE IT OUTSIDE OF A SECURE, TEST SETTING.             *
  ****************************************************************
.Link
Get-Process
#>

[cmdletbinding()]
Param()

Try {
  #test if the custom type has already been added
  [user32] -is [Type] | Out-Null
}
catch {
    #type not found so add it

    Add-Type -typeDefinition @"
        using System;
        using System.Runtime.InteropServices;

        public class User32
        {
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();
        }
"@
    #must left justify here-string closing @
} #catch

<#
get the process for the currently active foreground window as long as it has a value
greater than 0. A value of 0 typically means a non-interactive window. Also ignore
any Task Switch windows
#>

(Get-Process).where({$_.MainWindowHandle -eq ([user32]::GetForegroundWindow()) -AND $_.MainWindowHandle -ne 0 -AND $_.Name -ne 'Explorer' -AND $_.Title -notmatch "Task Switching"})


} #end Get-ForegroundWindowProcess

Function Get-WindowTime {

<#
.Synopsis
Monitor time by active window
.Description
This script will monitor how much time you spend based on how long a given window is active. Monitoring will continue until one of the specified triggers is detected. 

By default monitoring will continue for 1 minute. Use -Minutes to specify a different value. You can also specify a trigger by a specific date and time or by detection of a specific process.
.Parameter Time
Monitoring will continue until this datetime value is met or exceeded.
.Parameter Minutes
The numer of minutes to monitor. This is the default behavior.
.Parameter ProcessName
The name of a process that you would see with Get-Process, e.g. Notepad or Calc. Monitoring will stop when this process is detected.
Parameter AsJob
Run the monitoring in a background job. Note that if you stop the job you will NOT have any results.
.Example
PS C:\> $data = Get-WindowTime -minutes 60

Monitor window activity for the next 60 minutes. Be aware that you won't get your prompt back until this command completes.
.Example
PS C:\> Get-WindowTime -processname calc -asjob

Start monitoring windows in the background until the Calculator process is detected.
.Notes
Last Updated: October 7, 2015
Version     : 2.0

Learn more about PowerShell:
http://jdhitsolutions.com/blog/essential-powershell-resources/

  ****************************************************************
  * DO NOT USE IN A PRODUCTION ENVIRONMENT UNTIL YOU HAVE TESTED *
  * THOROUGHLY IN A LAB ENVIRONMENT. USE AT YOUR OWN RISK.  IF   *
  * YOU DO NOT UNDERSTAND WHAT THIS SCRIPT DOES OR HOW IT WORKS, *
  * DO NOT USE IT OUTSIDE OF A SECURE, TEST SETTING.             *
  ****************************************************************

.Link
Get-Process

#>

[cmdletbinding(DefaultParameterSetName= "Minutes")]
Param(
[Parameter(ParameterSetName="Time")]
[ValidateNotNullorEmpty()]
[DateTime]$Time,

[Parameter(ParameterSetName="Minutes")]
[ValidateScript({ $_ -ge 1})]
[Int]$Minutes = 1,

[Parameter(ParameterSetName="Process")]
[ValidateNotNullorEmpty()]
[string]$ProcessName,

[switch]$AsJob

)

Write-Verbose "[$(Get-Date)] Starting $($MyInvocation.Mycommand)"  

#define a scriptblock to use in the While loop
Switch ($PSCmdlet.ParameterSetName) {

"Time"    {    
            Write-Verbose "[$(Get-Date)] Stop monitoring at $Time"      
            [scriptblock]$Trigger = [scriptblock]::Create("(get-date) -ge ""$time""")
            Break
          }
"Minutes" {
            $Quit = (Get-Date).AddMinutes($Minutes)
            Write-Verbose "[$(Get-Date)] Stop monitoring in $minutes minute(s) at $Quit"
            [scriptblock]$Trigger = [scriptblock]::Create("(get-date) -ge ""$Quit""")
            Break  
          }
"Process" {
            If (Get-Process -name $processname -ErrorAction SilentlyContinue) {
                Write-Warning "The $ProcessName process is already running. Close it first then try again."
                #bail out
                Return
            }
            Write-Verbose "[$(Get-Date)] Stop monitoring after trigger $Processname"
           [scriptblock]$Trigger = [scriptblock]::Create("Get-Process -Name $ProcessName -ErrorAction SilentlyContinue")
            Break
          }

} #switch

#define the entire command as a scriptblock so it can be run as a job if necessary
$main = {
Param($sb)

if (-Not ($sb -is [scriptblock])) {
     #convert $sb to a scriptblock
     Write-Verbose "Creating sb from $sb"
     $sb = [scriptblock]::Create("$sb")
}

#create a hashtable
$hash=@{}

#create a collection of objects
$objs = @()

New-Variable -Name LastApp -Value $Null

while( -Not (&$sb) ) {

    $Process = Get-ForegroundWindowProcess

     [string]$app = $process.MainWindowTitle 

     if ( (-Not $app) -AND $process.MainModule.Description ) {
        #if no title but there is a description, use that
        $app = $process.MainModule.Description
     }
     elseif (-Not $app) {
        #otherwise use the module name
        $app = $process.mainmodule.modulename
     }
      
    if ($process -AND (($Process.MainWindowHandle -ne $LastProcess.MainWindowHandle) -OR ($app -ne $lastApp )) ) {
        Write-Verbose "[$(Get-Date)] NEW App changed to $app"
        
         #record $last
         if ($LastApp) {
          if ($objs.WindowTitle -contains $LastApp) {
              #update same app that was previously found
              Write-Verbose "[$(Get-Date)] updating existing object $LastApp"
              
              $existing = $objs | where { $_.WindowTitle -eq $LastApp }
 
              Write-Verbose "[$(Get-Date)] SW = $($sw.elapsed)"
                            
              $existing.Time+= $sw.Elapsed

              #include a detail property object
               
              $existing.Detail += [pscustomObject]@{
                StartTime = $start
                EndTime = Get-Date
                ProcessID = $lastProcess.ID
                Process = if ($LastProcess) {$LastProcess} else {$process}
               }
              Write-Verbose "[$(Get-Date)] Total time = $($existing.time)"
             }
            else {
                #create new object

                #include a detail property object
                [pscustomObject]$detail=@{
                  StartTime = $start
                  EndTime = Get-Date
                  ProcessID = $lastProcess.ID
                  Process = if ($LastProcess) {$LastProcess} else {$process}
                }
                Write-Verbose "[$(Get-Date)] Creating new object for $LastApp"
                Write-Verbose "[$(Get-Date)] Time = $($sw.elapsed)"

                #get categories
                $appCategory = (Select-XML -xml $MonitorCategories -xpath "//app[@name='$($LastMainModule.Description.Trim())']").node.parentnode.name
                
                if (!$appcategory) {
                    $appCategory = "None"
                }

                $obj = New-Object -TypeName PSobject -Property @{
                    WindowTitle = $LastApp
                    Application = $LastMainModule.Description #$LastProcess.MainModule.Description
                    Product = $LastMainModule.Product         #$LastProcess.MainModule.Product
                    Time = $sw.Elapsed
                    Detail = ,([pscustomObject]@{
                    StartTime = $start
                    EndTime = Get-Date
                    ProcessID = $lastProcess.ID
                    Process = if ($LastProcess) {$LastProcess} else {$process}
                    } )
                    Category = $appCategory
                    Computername = $env:COMPUTERNAME
                    
                    }    

                $obj.psobject.TypeNames.Insert(0,"My.Monitored.Window")
                #add a custom type name

                #add the object to the collection
                $objs += $obj
          } #else create new object
        } #if $lastApp was defined
         else {
        Write-Verbose "You should only see this once"
        }

        #new Process with a window
        Write-Verbose "[$(Get-Date)] Start a timer"
        $SW = [System.Diagnostics.Stopwatch]::StartNew()     
        $start = Get-Date
  
        #set the last app
        $LastApp = $app
        #preserve process information
        $LastProcess = $Process
        $LastMainModule = $process.mainmodule

      #clear app just in case
      Remove-Variable app
  }
      Start-Sleep -Milliseconds 100

} #while

#update last app
if ($objs.WindowTitle -contains $LastApp) {
    #update same app that was previously found
    Write-Verbose "[$(Get-Date)] processing last object"
    Write-Verbose "[$(Get-Date)] updating existing object for $LastApp"
    Write-Verbose "[$(Get-Date)] SW = $($sw.elapsed)"
    $existing = $objs | where { $_.WindowTitle -eq $LastApp }

    $existing.Time+= $sw.Elapsed

    Write-Verbose "[$(Get-Date)] Total time = $($existing.time)"

    #include a detail property object
    
    $existing.Detail += [pscustomObject]@{
        StartTime = $start
        EndTime = Get-Date
        ProcessID = $lastProcess.ID
        Process = if ($LastProcess) {$LastProcess} else {$process}
    }
}
else {
    #create new object

    Write-Verbose "[$(Get-Date)] Creating new object"
    Write-Verbose "[$(Get-Date)] Time = $($sw.elapsed)"

    #get categories
    $appCategory = (Select-XML -xml $MonitorCategories -xpath "//app[@name='$($LastMainModule.Description.Trim())']").node.parentnode.name
                
    if (!$appcategory) {
        $appCategory = "None"
    }

    $obj = New-Object -TypeName PSobject -Property @{
        WindowTitle = $LastApp
        Application = $LastMainModule.Description #$LastProcess.MainModule.Description
        Product = $LastMainModule.Product         #$LastProcess.MainModule.Product
        Time = $sw.Elapsed
        Detail = ,([pscustomObject]@{
        StartTime = $start
        EndTime = Get-Date
        ProcessID = $lastProcess.ID
        Process = if ($LastProcess) {$LastProcess} else {$process}
         })
        Category = $appCategory
        Computername = $env:COMPUTERNAME
        }    

    $obj.psobject.TypeNames.Insert(0,"My.Monitored.Window")
    #add a custom type name

    #add the object to the collection
    $objs += $obj
} #else create new object

$objs

Write-Verbose "[$(Get-Date)] Ending $($MyInvocation.Mycommand)"  
} #main


if ($asJob) {
    Write-Verbose "Running as background job"
   Start-Job -ScriptBlock $main -ArgumentList @($Trigger) -InitializationScript  {Import-Module MyMonitor}

}
else {
    #run it
   Invoke-Command -ScriptBlock $main -ArgumentList @($Trigger)
}

} #end Get-WindowTime

Function Measure-WindowTotal {

<#
.Synopsis
Measure Window usage results.
.Description
This command is designed to take output from Get-WindowTime and measure total time either by Application, the default, by Product or Window title. Or you can elect to get the total time for all piped in measured window objects.

You can also filter based on keywords found in the window title. See examples.
.Parameter Filter
A string to be used for filtering based on the window title. The string can be a regular expression pattern.
.Example
PS C:\> $data = Get-WindowTime -ProcessName calc  

Start monitoring windows until calc is detected as a running process.

PS C:\> $data | Measure-WindowTotal

Application                                                        TotalTime
-----------                                                        ---------
Desktop Window Manager                                             00:00:05.5737147
Dropbox                                                            00:00:05.9456759
Skype                                                              00:00:07.3145639
Microsoft Management Console                                       00:00:08.4824010
COM Surrogate                                                      00:00:08.6499505
SugarSync                                                          00:00:11.3705429
Wireshark                                                          00:00:21.0674948
Windows PowerShell                                                 00:00:33.6266261
Virtual Machine Connection                                         00:02:07.3252447
Spotify                                                            00:04:25.0623684
Thunderbird                                                        00:17:38.3666410
Windows PowerShell ISE                                             00:20:28.3669673
Microsoft Word                                                     00:22:18.5841682
Waterfox                                                           01:20:46.0559866


PS C:\> $data  | Measure-WindowTotal -Product

Product                                                            TotalTime
-------                                                            ---------
Dropbox                                                            00:00:05.9456759
Skype                                                              00:00:07.3145639
SugarSync                                                          00:00:11.3705429
Wireshark                                                          00:00:21.0674948
Spotify                                                            00:04:25.0623684
Thunderbird                                                        00:17:38.3666410
Microsoft Office 2013                                              00:22:18.5841682
Microsoft® Windows® Operating System                               00:23:32.0249043
Waterfox                                                           01:20:46.0559866

The first command gets data from active window usage. The second command measures the results by Application. The last command measures the same data but by the product property.
.Example
PS C:\> $data | Measure-WindowTotal -filter "facebook" -TimeOnly

Days              : 0
Hours             : 0
Minutes           : 11
Seconds           : 2
Milliseconds      : 781
Ticks             : 6627813559
TotalDays         : 0.00767108050810185
TotalHours        : 0.184105932194444
TotalMinutes      : 11.0463559316667
TotalSeconds      : 662.7813559
TotalMilliseconds : 662781.3559

Get just the time that was spent on any window that had Facebook in the title.
.Example
PS C:\> $data | Measure-WindowTotal -filter "facebook|twitter"

Application                                                        TotalTime
-----------                                                        ---------
Desktop Window Manager                                             00:00:05.5737147
Waterfox                                                           00:10:57.2076412

Display how much time was spent on Facebook or Twitter.

.Example
PS C:\> Measure-WindowTotal $data -Category | Sort Time -descending | format-table -AutoSize

Category      Count Time
--------      ----- ----
Internet         68 01:25:18.4329189
Office            5 00:22:18.5841682
PowerShell        4 00:21:01.9935934
Test              1 00:20:28.3669673
Development       1 00:20:28.3669673
Mail             29 00:17:38.3666410
None              5 00:02:27.4945858
Utilities         2 00:00:29.5498958
Cloud             1 00:00:11.3705429
Communication     1 00:00:07.3145639

Measure window totals by category then sort the results by time in descending order. The result is formatted as a table.

.Notes
Last Updated: October 7, 2015
Version     : 2.0

Learn more about PowerShell:
http://jdhitsolutions.com/blog/essential-powershell-resources/

  ****************************************************************
  * DO NOT USE IN A PRODUCTION ENVIRONMENT UNTIL YOU HAVE TESTED *
  * THOROUGHLY IN A LAB ENVIRONMENT. USE AT YOUR OWN RISK.  IF   *
  * YOU DO NOT UNDERSTAND WHAT THIS SCRIPT DOES OR HOW IT WORKS, *
  * DO NOT USE IT OUTSIDE OF A SECURE, TEST SETTING.             *
  ****************************************************************

.Link
Get-WindowTime
#>

[cmdletbinding(DefaultParameterSetName="Product")]
Param(
[Parameter(Position=0,ValueFromPipeline)]
[Parameter(ParameterSetName="Product")]
[Parameter(ParameterSetName="Title")]
[Parameter(ParameterSetName="Category")]
[ValidateNotNullorEmpty()]
$InputObject,
[Parameter(ParameterSetName="Product")]
[Switch]$Product,
[Parameter(ParameterSetName="Title")]
[Switch]$Title,
[Parameter(ParameterSetName="Category")]
[Switch]$Category,
[Parameter(ParameterSetName="Product")]
[Parameter(ParameterSetName="Title")]
[ValidateNotNullorEmpty()]
[String]$Filter=".*",
[Parameter(ParameterSetName="Product")]
[Parameter(ParameterSetName="Title")]
[Switch]$TimeOnly
)

Begin {
    Write-Verbose "Starting $($MyInvocation.Mycommand)"  

    #initialize
    $hash=@{}

    if ($Product) {
      $objFilter = "Product"
    }
    elseif($Title) {
      $objFilter = "WindowTitle"
    }
    elseif($Category) {
      $objFilter = "Category"
      #initialize an array to hold incoming items
      $data=@()
    }
    else {
      $objFilter = "Application"
    }

    Write-Verbose "Calculating totals by $objFilter"
} #begin

Process {

foreach ($item in $InputObject) {
  #only process objects where the window title matches the filter which
  #by default is everything and there is only one object in the product
  #which should eliminate task switching data
  if ($item.WindowTitle -match $filter -AND $item.Product.count -eq 1) {

    If ($Category) {
        $data+=$item
    }
    else {
      if ($hash.ContainsKey($item.$objFilter)) { 
        #update an existing entry in the hash table
         $hash.Item($item.$objFilter) += $item.Time 
      }
      else {
        #Add an entry to the hashtable
        Write-Verbose "Adding $($item.$objFilter)"
        $hash.Add($item.$objFilter,$item.time)
      }
    } #else not -Category
  }
} #foreach
} #process

End {
    Write-Verbose "Processing data"

    if ($Category) {
        Write-Verbose "Getting category breakdown"
        $output = $data.category | select -Unique | foreach {
         $thecategory = $_
         $hash = [ordered]@{Category=$theCategory}
         $items = $($data).Where({$_.category -contains $thecategory})
         $totaltime = $items | foreach -begin {$time = new-timespan} -process {$time+=$_.time} -end {$time}
         $hash.Add("Count",$items.count)
         $hash.Add("Time",$TotalTime)
         [pscustomobject]$hash
         }
    }
    else {
        #turn hash table into a custom object and sort on time by default
          $output = ($hash.GetEnumerator()).foreach({
          [pscustomobject]@{$objfilter=$_.Name;"TotalTime"=$_.Value}
          
        }) | Sort TotalTime
    }

    if ($TimeOnly) {
        $output | foreach -begin {$total = New-TimeSpan} -process {$total+=$_.Totaltime} -end {$total}
     }
     else {
        $output # | Select $objFilter,TotalTime
     }

      Write-Verbose "Ending $($MyInvocation.Mycommand)"
} #end

} #end Measure-WindowTotal

Function Get-WindowTimeSummary {
<#
.Synopsis
Get a summary of window usage time
.Description
This command will take an array of window usage data and present a summary based on application. The output will include the total time as well as the first and last times for that particular application.

As an alternative you can get a summary by Product or you can filter using a regular expression pattern on the window title.
.Example
PS C:> 

PS C:\> Get-WindowTimeSummary $data

Name                              Total                        Start                            End
----                              -----                        -----                            ---
Windows PowerShell ISE            00:20:28.3669673             10/7/2015 8:07:09 AM             10/7/2015 10:36:36 AM
SugarSync                         00:00:11.3705429             10/7/2015 8:07:20 AM             10/7/2015 9:01:31 AM
Spotify                           00:04:25.0623684             10/7/2015 8:07:34 AM             10/7/2015 8:39:27 AM
Thunderbird                       00:17:38.3666410             10/7/2015 10:24:47 AM            10/7/2015 10:30:52 AM
COM Surrogate                     00:00:08.6499505             10/7/2015 8:09:51 AM             10/7/2015 8:10:00 AM
Waterfox                          01:20:46.0559866             10/7/2015 9:19:33 AM             10/7/2015 10:30:54 AM
Wireshark                         00:00:21.0674948             10/7/2015 8:13:23 AM             10/7/2015 8:29:48 AM
Virtual Machine Connection        00:02:07.3252447             10/7/2015 10:12:46 AM            10/7/2015 10:12:43 AM
Skype                             00:00:07.3145639             10/7/2015 8:29:27 AM             10/7/2015 8:39:12 AM
Windows PowerShell                00:00:33.6266261             10/7/2015 8:30:59 AM             10/7/2015 8:40:09 AM
Dropbox                           00:00:05.9456759             10/7/2015 8:32:26 AM             10/7/2015 8:32:32 AM
Microsoft Management Console      00:00:08.4824010             10/7/2015 8:39:45 AM             10/7/2015 10:12:46 AM
Microsoft Word                    00:22:18.5841682             10/7/2015 9:04:00 AM             10/7/2015 10:04:05 AM
Desktop Window Manager            00:00:05.5737147             10/7/2015 9:19:57 AM             10/7/2015 9:20:02 AM

Get time summary using the default application type.

.Example
PS C:\> Get-WindowTimeSummary $data -Type Product

Name                              Total                        Start                            End
----                              -----                        -----                            ---
Microsoft® Windows® Operating ... 00:23:32.0249043             10/7/2015 8:09:51 AM             10/7/2015 10:12:43 AM
SugarSync                         00:00:11.3705429             10/7/2015 8:07:20 AM             10/7/2015 9:01:31 AM
Spotify                           00:04:25.0623684             10/7/2015 8:07:34 AM             10/7/2015 8:39:27 AM
Thunderbird                       00:17:38.3666410             10/7/2015 10:24:47 AM            10/7/2015 10:30:52 AM
Waterfox                          01:20:46.0559866             10/7/2015 9:19:33 AM             10/7/2015 10:30:54 AM
Wireshark                         00:00:21.0674948             10/7/2015 8:13:23 AM             10/7/2015 8:29:48 AM
Skype                             00:00:07.3145639             10/7/2015 8:29:27 AM             10/7/2015 8:39:12 AM
Dropbox                           00:00:05.9456759             10/7/2015 8:32:26 AM             10/7/2015 8:32:32 AM
Microsoft Office 2013             00:22:18.5841682             10/7/2015 9:04:00 AM             10/7/2015 10:04:05 AM

Get time summary by product.
.Example
PS C:\> Get-WindowTimeSummary $data -filter "facebook|hootsuite"

Name                              Total                        Start                            End
----                              -----                        -----                            ---
facebook|hootsuite                00:28:22.3692839             10/7/2015 8:33:47 AM             10/7/2015 10:30:54 AM

Filter window titles with a regular expression.
.Notes
Last Updated: October 7, 2015
Version     : 2.0

Learn more about PowerShell:
http://jdhitsolutions.com/blog/essential-powershell-resources/

  ****************************************************************
  * DO NOT USE IN A PRODUCTION ENVIRONMENT UNTIL YOU HAVE TESTED *
  * THOROUGHLY IN A LAB ENVIRONMENT. USE AT YOUR OWN RISK.  IF   *
  * YOU DO NOT UNDERSTAND WHAT THIS SCRIPT DOES OR HOW IT WORKS, *
  * DO NOT USE IT OUTSIDE OF A SECURE, TEST SETTING.             *
  ****************************************************************

.Link
Get-WindowTime
Measure-WindowTotal
#>

[cmdletbinding(DefaultParameterSetName="Type")]
Param(
[Parameter(
Position=0,Mandatory,
HelpMessage="Enter a variable with your Window usage data")]
[ValidateNotNullorEmpty()]
$Data,
[Parameter(ParameterSetName="Type")]
[ValidateSet("Product","Application")]
[string]$Type="Application",

[Parameter(ParameterSetName="Filter")]
[ValidateNotNullorEmpty()]
[string]$Filter

)

Write-Verbose -Message "Starting $($MyInvocation.Mycommand)"  


if ($PSCmdlet.ParameterSetName -eq 'Type') {
    Write-Verbose "Processing on $Type"
    #filter out blanks and objects with multiple products from ALT-Tabbing
    $grouped = ($data).Where({$_.$Type -AND $_.$Type.Count -eq 1}) | Group-Object -Property $Type
}
else {
  #use filter
  Write-Verbose "Processing on filter: $Filter"
  $grouped = ($data).where({$_.WindowTitle -match $Filter -AND $_.Product.Count -eq 1}) |
  Group-Object -Property {$Filter}
}

if ($Grouped) {
    $grouped| Select Name,
    @{Name="Total";Expression={ 
    $_.Group | foreach -begin {$total = New-TimeSpan} -process {$total+=$_.time} -end {$total}
    }},
    @{Name="Start";Expression={
    ($_.group | sort Detail).Detail[0].StartTime
    }},
    @{Name="End";Expression={
    ($_.group | sort Detail).Detail[-1].EndTime
    }}
}
else {
    Write-Warning "No items found"
  }

    Write-Verbose -Message "Ending $($MyInvocation.Mycommand)"

} #end Get-WindowsTimeSummary

#endregion

#region TypeData

#set default display property set
Update-TypeData -TypeName "my.monitored.window" -DefaultDisplayPropertySet "Time","Application","WindowTitle","Product" -DefaultDisplayProperty WindowTitle -Force
Update-TypeData -TypeName "deserialized.my.monitored.window" -DefaultDisplayPropertySet "Time","Application","WindowTitle","Product" -DefaultDisplayProperty WindowTitle  -Force

#add an alias for the WindowTitle property
Update-TypeData -TypeName "My.Monitored.Window" -MemberType AliasProperty -MemberName Title -Value WindowTitle -force
Update-TypeData -TypeName "deserialized.My.Monitored.Window" -MemberType AliasProperty -MemberName Title -Value WindowTitle -force

#endregion

#region Aliases

Set-Alias -name mwt -Value Measure-WindowTotal
Set-Alias -name gfwp -Value Get-ForegroundWindowProcess
Set-Alias -Name gwt -Value Get-WindowTime
Set-Alias -name gwts -Value Get-WindowTimeSummary

#endregion

#region Variables

<#
Look for a copy of Categories.xml in the user's PowerShell directory and use that if found.
Otherwise use the one included with the module. This is to prevent overwriting the xml
file in future module updates.
#>

$localXML = Join-Path -Path $env:USERPROFILE -ChildPath "Documents\WindowsPowerShell\Categories.xml"
$modXML = Join-Path -path $PSScriptroot -ChildPath categories.xml

if (Test-Path -path $localXML) {
    [xml]$MonitorCategories = Get-Content -Path $localXML
} 
else {
    [xml]$MonitorCategories = Get-Content -Path $modXML
}


#endregion

Export-ModuleMember -Function * -Alias * -Variable MonitorCategories

