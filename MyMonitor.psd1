<#
last updated 11 November 2015

  ****************************************************************
  * DO NOT USE IN A PRODUCTION ENVIRONMENT UNTIL YOU HAVE TESTED *
  * THOROUGHLY IN A LAB ENVIRONMENT. USE AT YOUR OWN RISK.  IF   *
  * YOU DO NOT UNDERSTAND WHAT THIS SCRIPT DOES OR HOW IT WORKS, *
  * DO NOT USE IT OUTSIDE OF A SECURE, TEST SETTING.             *
  ****************************************************************
#>

@{
Author="Jeff Hicks [jhicks@jdhitsolutions.com]"
CmdletsToExport=@()
CompanyName="JDH Information Technology Solutions, Inc."
Copyright="2010-2016 © JDH Information Technology Solutions, Inc."
Description="Functions to track and measure active window usage"
CLRVersion="3.0"
FileList=@()
FormatsToProcess=@("MyMonitor.format.ps1xml")
FunctionsToExport=@("Get-ForegroundWindowProcess","Get-WindowTime",
"Get-WindowTimeSummary","Measure-WindowTotal")
AliasesToExport=@("gfwp","gwt","mwt","gwts")
GUID="7a4bd35a-7acf-4cdb-b929-069dd86a0525"
RootModule ="MyMonitor.psm1"
ModuleVersion ="2.1"
PowerShellVersion ="4.0"
PrivateData = @{
 PSData = @{

        # Tags applied to this module. These help with module discovery in online galleries.
        Tags = @("Monitoring","TimeTracking")

        # A URL to the license for this module.
        LicenseUri = 'https://github.com/jdhitsolutions/MyMonitor/blob/master/License.txt'

        # A URL to the main website for this project.
        ProjectUri = 'https://github.com/jdhitsolutions/MyMonitor'

        # A URL to an icon representing this module.
        # IconUri = ''

        # ReleaseNotes of this module
        # ReleaseNotes = ''

    } # End of PSData hashtable
    }
RequiredAssemblies = @()
RequiredModules = @()
ScriptsToProcess = @()
TypesToProcess = @()
VariablesToExport = "MonitorCategories"
}