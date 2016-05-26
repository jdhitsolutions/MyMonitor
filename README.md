#MyMonitor

The MyMonitor module contains a small set of commands for monitoring active 
window use. The Get-WindowTime function uses another "helper" function called 
Get-ForegroundWindowProcess to retrieve the active window process. However, 
processes that don't have interactive windows or those associated with Task 
Switching are excluded. 
    
Get-WindowTime keeps track of how long you are spending on a window based 
primarily on its window title. At the end of the monitoring process, it writes
results to the pipeline.

    Time        : 00:00:18.9084682
    Application : Thunderbird
    WindowTitle : Junk - Jeffery Hicks - Mozilla Thunderbird
    Product     : Thunderbird

    Time        : 00:00:12.6757123
    Application : Windows PowerShell ISE
    WindowTitle : Administrator: Windows PowerShell ISE
    Product     : Microsoft® Windows® Operating System

##MONITORING
By default, Get-WindowTime will monitor active window use for 1 minute. You can
specify a different number of minute or a specific date time. Another option is
to use a process "trigger". This is the name of a process that you rarely use, 
such as CALC. When Get-WindowTime detects the presence of that process, it will 
stop monitoring and write results to the pipeline. 
    
Because monitoring will block the prompt, you also have an option to run 
everything in a background job. Note that if you terminate the background job 
you will not get any results.

One potential drawback to this toolset, it that there is no way to pause and 
resume monitoring. Thus, you might have skewed results when you go to meetings 
or take a lunch break. One workaround is to select an application that you do 
not use, such as Calc.exe and run that during your away time. When you measure
results, you can ignore or filter out that process.

##MEASURING
When you are finished monitoring, you can use Measure-WindowTotal to analyze 
the data. Pipe your data to the command. The default is to display results by 
application.

    PS C:\> $data | Measure-WindowTotal

    Application                                       TotalTime
    -----------                                       ---------
    Windows PowerShell ISE                            00:00:12.6757123
    Windows PowerShell                                00:00:22.6393253
    Google Chrome                                     00:01:54.1870428
    Thunderbird                                       00:03:10.1524340
    Microsoft Management Console                      00:04:19.7252959

You can also measure by product:

    PS C:\> $data | Measure-WindowTotal -Product

    Product                                           TotalTime
    -------                                           ---------
    Google Chrome                                     00:01:54.1870428
    Thunderbird                                       00:03:10.1524340
    Microsoft® Windows® Operating System              00:04:55.0403335

or by Title:

    PS C:\> $data | Measure-WindowTotal -Title

    WindowTitle                                       TotalTime
    -----------                                       ---------
    July 2015 - Mozilla Thunderbird                   00:00:00.7309264
    New Tab - Google Chrome                           00:00:02.6440093
    Calendar - Mozilla Thunderbird                    00:00:02.7084472
    Hootsuite - Google Chrome                         00:00:07.6023926
    Facebook - Google Chrome                          00:00:09.0954287
    Administrator: Windows PowerShell ISE             00:00:12.6757123
    Junk - Jeffery Hicks - Mozilla Thunderbird        00:00:18.9084682
    Administrator: Windows PowerShell                 00:00:22.6393253
    Inbox - Jeffery Hicks - Mozilla Thunderbird       00:00:29.1373713
    Hyper-V Manager                                   00:00:49.1406603
    Update Services                                   00:03:23.7454981

You can also filter by specifying a string or regular expression pattern for 
text in the Window title. This can be useful when you want to see how much time
you are spending on a particular set of applications or products.

    PS C:\> $data | Measure-WindowTotal -Filter "facebook|hootsuite"

    Application                                       TotalTime
    -----------                                       ---------
    Google Chrome                                     00:00:16.6978213

The last measure-related command is Get-WindowTimeSummary. To use this you will 
need a variable holding all of the work items. This will show you not only the 
total time for products or applications, but also a range of the first and last
item for each.

    PS C:\> get-windowtimesummary $data | sort total | format-table -autosize

    Name                                 Total            Start                 End
    ----                                 -----            -----                 ---
    Google Chrome Windows PowerShell ISE 00:00:00.2776561 8/25/2015 10:46:36 AM 8/25/2015 11:01:04 AM
    Microsoft Word                       00:00:01.4462946 8/25/2015 12:14:50 PM 8/25/2015 12:18:34 PM
    Windows PowerShell                   00:04:50.0923401 8/25/2015 11:38:14 AM 8/25/2015 12:19:31 PM
    Thunderbird                          00:05:35.3980271 8/25/2015 12:02:32 PM 8/25/2015 12:19:04 PM
    Google Chrome                        00:29:57.3321071 8/25/2015 11:56:36 AM 8/25/2015 12:20:27 PM

##CATEGORIES
The module now includes support for categories. You can assign a category to 
an application. The application is the value of the Description property you 
see with Get-Process.

    PS C:\> get-process winword | select name,description

    Name                                                               Description                                                       
    ----                                                               -----------                                                       
    WINWORD                                                            Microsoft Word

In this situation, you can assign a category to "Microsoft Word". Categories 
are stored in a special XML file called Categories.xml. A sample entry looks like this:

     <category>
        <name>Internet</name>
            <app name = "Waterfox"/>
            <app name = "Firefox"/>
            <app name = "Internet Explorer"/>
            <app name = "Google Chrome"/>
            <app name = "Skype"/>
            <app name = "Spotify"/>
            <app name = "Amazon Music"/>
    </category>

You can assign an application to multiple categories. Note that the value you 
enter for the name is case sensitive. 
    
The module includes a "starter" version of Categories.xml. It is strongly 
recommended that you create a copy in your Windows PowerShell directory using 
the same name. This is to avoid overwriting your customizations during module 
updates. You can edit this file as much as you would like. 

When you import the module, it will first look for a copy of Categories.xml in 
your PowerShell directory and use it if found. Otherwise, it will use the file
in the module script root.

After you have gathered data, you can use the categories.

    PS C:\> measure-windowtotal $data -Category | format-table -AutoSize

    Category      Count Time            
    --------      ----- ----            
    PowerShell        4 00:21:01.9935934
    Test              1 00:20:28.3669673
    Development       1 00:20:28.3669673
    Cloud             1 00:00:11.3705429
    Internet         68 01:25:18.4329189
    Mail             29 00:17:38.3666410
    None              5 00:02:27.4945858
    Utilities         2 00:00:29.5498958
    Communication     1 00:00:07.3145639
    Office            5 00:22:18.5841682

Or like this:

    PS C:\> ($data).Where({$_.Category -contains "Internet"}) | Measure-WindowTotal

    Application                                        TotalTime   
    -----------                                        ---------                                                         
    Skype                                       00:00:07.3145639
    Spotify                                     00:04:25.0623684
    Waterfox                                    01:20:46.0559866

##CUSTOM VIEWS
The MyMonitor module includes custom type and format extensions. The default 
display is a list. The module includes several views for Format-Table, but you 
should sort your data first, as the views group data by either product:
    
    PS C:\> $data | sort Product | format-table -View Product

or application:

    PS C:\> $data | sort Application,Time | format-table -View Application
    
The output from Get-WindowTime is no different than any other type of object 
in PowerShell which means you can export to a CSV or XML format, save to a 
text file or anything else. These commands only work for the current and 
local interactive user.

##ALIASES
The module contains the following aliases:

    gfwp --> Get-ForegroundWindowProcess
    gwt  --> Get-Windowtime
    mwt  --> Measure-WindowTotal
    gwts --> Get-WindowTimeSummary

##IMPORTING AND EXPORTING
It is recommended that you use the XML format to export and import work history.
    
    PS C:\> $data | export-clixml C:\work\8-25-15-work.xml

The XML format will preserve property type. Before using any imported data, be 
sure to import this module first. This will properly format and display the 
imported data and allow you to measure it.
    
##KNOWN ISSUES
Occassionally, you may see a result like this:
    
    Time        : 00:00:00.2095710
    Application : {Google Chrome, Thunderbird}
    WindowTitle : Hootsuite - Google Chrome Mozilla Thunderbird

This appears to be the result of task switching with either Alt+Tab or Ctrl+
Tab. For the most part, the time spent is in milliseconds and thus can be 
ignored. 