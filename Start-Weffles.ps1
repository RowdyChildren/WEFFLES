<#
.Synopsis
   Monitors event forwarding logs and generates a csv that can be used by a PowerBI
.DESCRIPTION
   Monitors event forwarding logs and generates a csv that can be used by a PowerBI
.PARAMETER OutputPath
    If set this is the fully qualified path to the output file
.PARAMETER DataDirectory
    If set this will put the Data directory somerwhere other than the same directory as the script
.EXAMPLE
   This outputs the csv to a custom location:
    .\Start-Weffles.ps1 -OutputFile "C:\customlocation\weffles.csv"
.EXAMPLE
   This sets the data directory to somewhere other than the same directory as Start-Weffles.ps1
    .\Start-Weffles.ps1 -WorkingDirectory "C:\customlocation"
.EXAMPLE
    This starts Weffles and uses the same directory the script is in as the data directory
    .\Start-Weffles.ps1
.Link
    https://aka.ms/weffles
    https://aka.ms/jepayne
#>
#Requires -Version 5
[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [String]
    $DataDirectory = $PSScriptRoot,
    [Parameter(Mandatory=$false)]
    [string]
    $OutputPath = "$DataDirectory\weffles.csv"
)
try {
    Import-Module .\EventLogWatcher.psm1
} catch {
    Write-Error "Unable to import .\EventLogWatcher.psm1 due to $($Error[0])"
}

#Classes
class Config {
    [string]$EventLogName
    [int[]]$EventIdentifiers
    [Hashtable]$CustomEventFields
}

function Get-EventLogWatcher() {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True)]
        [Int[]]
        $EventIdentifiers,
        [Parameter(Mandatory=$True)]
        [string]
        $EventLogName
    )
    #Get the existing bookmark, if there is one
    if (Test-Path "$DataDirectory\bookmark.stream") {
        $_Bookmark = ((Get-Content "$DataDirectory\bookmark.stream")[1])
        $_EventLogName = ($_Bookmark -split "'")[1]
        $_EventRecordId = ($_Bookmark -split "'")[3]
        #Test if the bookmark is still valid
        if (!(Get-WinEvent -LogName $_EventLogName -FilterXPath "*[System[(EventRecordID=$_EventRecordId)]]") -or ($EventLogName -ne $_EventLogName)) {
            Remove-Item "$DataDirectory\bookmark.stream"
        }
    }
    #Lets get the bookmark, if its valid it will be there, if it isn't we create a new one
    $_BookmarkToStartFrom = Get-BookmarkToStartFrom -BookmarkStreamPath "$DataDirectory\bookmark.stream"
    
    #Build the event query
    $_Query = ""
    foreach ($EventIdentifier in $EventIdentifiers) {
        $_Query = "*[System[EventID=$EventIdentifier]] or $_Query"
    }

    $_EventLogQuery = New-EventLogQuery $EventLogName -Query $_Query
    return New-EventLogWatcher -EventLogQuery $_EventLogQuery -BookmarkToStartFrom $_BookmarkToStartFrom
}

function Get-Config() {
    $_ConfigFile = Get-Content -Raw -Path "$DataDirectory\config.json" | ConvertFrom-Json -AsHashtable
    $_Config =  [Config]::new()
    $_Config.EventLogName = $_ConfigFile["EventLogName"]
    $_Config.EventIdentifiers = $_ConfigFile["EventIdentifiers"]
    $_Config.CustomEventFields = $_ConfigFile["CustomEventLogFields"]
    return $_Config
}

#Following added to unregister any existing watcher Event in case script has already been run
if (Get-Event -SourceIdentifier "Weffles") {
    try {
        Remove-Event -SourceIdentifier "Weffles"
    } catch {
        throw $Error[0]
        exit 1
    }
}

[Config]$config = Get-Config

$EventLogWatcher = Get-EventLogWatcher -EventIdentifiers $config.EventIdentifiers -EventLogName $config.EventLogName

$EventAction = {
    #Creating object to output to .csv
    $theEvent = New-Object [PSCustomObject]
    #Add Default Event Fields 
    $theEvent | Add-Member NoteProperty EventDate $EventRecord.TimeCreated.ToShortDateString()
    $theEvent | Add-Member NoteProperty EventHost $EventRecord.MachineName
    $theEvent | Add-Member NoteProperty EventID $EventRecord.ID
    #itterate thorugh configuration file and build the object
    foreach ($key in $config.CustomEventFields.Keys) {
        $theEvent | Add-Member NoteProperty $key $EventRecordXml.SelectSingleNode("//*[@Name='$($configuration.CustomEventFields[$key])']")."#text"
    }
	#Adding UserID at the end
    $theEvent | Add-Member NoteProperty UserID $EventRecord.UserID
    #Convert the object of CSV and skip the first row to avoid the headers
    $theEvent | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 | Out-File $OutputPath -Encoding default -Append
}

#Register the watcher
Register-EventRecordWrittenEvent $EventLogWatcher -action $EventAction -SourceIdentifier "Weffles"
#Enable the Watcher
$EventLogWatcher.Enabled = $True