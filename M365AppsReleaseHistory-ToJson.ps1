<#

.Version 1.0.0

.Author Adrian Cojocaru

.Date 10-Dec-2024

.Synopsis
    Convert the Microsoft 365 Apps version history HTML table to Json.

.Description
    Convert https://learn.microsoft.com/en-us/officeupdates/update-history-microsoft365-apps-by-date#version-history to Json with the help of the HTMLAgilityPack module
    The json file is used to determine how old is the O365 a system is running.
    Since O365 does not write the App version on the machine that is installed on, this script will also utilize the Version History table to create a JSON for matching versions to builds
    The resulting Json can be uploaded on a blob storage so it can be read from the client.

#>
[string]$Global:SASToken = if ($env:AZUREPS_HOST_ENVIRONMENT -or $PSPrivateMetadata.JobId) { Get-AutomationVariable -Name "AzFileUploadSASToken" }
[string]$Global:StorageAccountName = if ($env:AZUREPS_HOST_ENVIRONMENT -or $PSPrivateMetadata.JobId) { Get-AutomationVariable -Name "StorageAccountName" }
#Region ----------------------------------------------------- [Classes] ----------------------------------------------
class CustomException : Exception {
    <#

    .DESCRIPTION
    Used to throw exceptions.
    .EXAMPLE
    throw [CustomException]::new( "Get-ErrorOne", "This will cause the script to end with ExitCode 101")

#>
    [string] $additionalData

    CustomException($Message, $additionalData) : base($Message) {
        $this.additionalData = $additionalData
    }
}
#EndRegion -------------------------------------------------- [Classes] ----------------------------------------------

#Region ----------------------------------------------------- [Functions] ----------------------------------------------
Function Write-LogRunbook {
    <#

    .DESCRIPTION
    Write messages to a log file defined by $LogPath and also display them in the console.
    Message format: [Date & Time] [CallerInfo] :: Message Text

#>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [ValidateNotNull()]
        [AllowEmptyString()]
        # Mandatory. Specifies the message string.
        [string]$Message,
        [Parameter(Mandatory = $false, Position = 1)]
        [ValidateNotNull()]
        # Optional. Specifies the name of the message writter. Function, command or custom name. Defaults to FunctioName or unknown
        [string]$Caller = 'Unknown'
    )
    Begin {
        [string]$LogDateTime = (Get-Date -Format 'MM-dd-yyyy HH\:mm\:ss.fff').ToString()
    }
    Process {
        "[$LogDateTime] [${Caller}] :: $Message" | Write-Verbose -Verbose  
    }
    End {}
}

function Write-ErrorRunbook {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $false, Position = 0)]
        [AllowEmptyCollection()]
        # Optional. The errorr collection.
        [array]$ErrorRecord
    )
    Begin {
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        If (-not $ErrorRecord) {
            If ($global:Error.Count -eq 0) {
                Return
            }
            Else {
                [array]$ErrorRecord = $global:Error[0]
            }
        }
    }
    Process {
        [string]$LogDateTime = (Get-Date -Format 'MM-dd-yyyy HH\:mm\:ss.fff').ToString()
        $ErrorRecord | ForEach-Object {
            $errNumber = $ErrorRecord.count - $( $ErrorRecord.IndexOf($_))
            if ($_.Exception.GetType().Name -eq 'CustomException') {
                $ErrorText = "[$LogDateTime] [${CmdletName} Nr. $errNumber] " + `
                    "Line: $($($_.InvocationInfo).ScriptLineNumber) Char: $($($_.InvocationInfo).OffsetInLine) " + `
                    "[$($($_.Exception).Message)] $($($_.Exception).additionalData)" 
            }
            else {
                $ErrorText = "[$LogDateTime] [${CmdletName} Nr. $errNumber] :: $($($_.Exception).Message)`n" + `
                    ">>> Line: $($($_.InvocationInfo).ScriptLineNumber) Char: $($($_.InvocationInfo).OffsetInLine) <<<`n" + `
                    "$($($_.InvocationInfo).Line)"
                if ($ErrorRecord.ErrorDetails.Message) {
                    $ErrorText += $ErrorRecord.ErrorDetails.Message
                }
            }
            $ErrorText | Write-Error
        }
    }
    End {}
}
function Add-FileToBlobStorage {
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path $_ })]
        [string]$FilePath
        #[Parameter(Mandatory = $false)]
        #[ValidateScript({ $_ -match "https\:\/\/(.)*\.blob.core.windows.net\/(.)*\?(.)*" })]
        #$uri = "https://$StorageAccountName.blob.core.windows.net/dwc/$($name)$SASToken"
    )
    try {
        [string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
        $PSBoundParameters.GetEnumerator() | Sort-Object -Property Name | ForEach-Object { "$($_.Key) = $($_.Value)" | Write-LogRunbook -Caller $CmdletName }
        $FileName = (Get-Item $FilePath).Name
        [string]$uri = "https://$StorageAccountName.blob.core.windows.net/dwc/$($FileName)$SASToken"
        $headers = @{
            'x-ms-blob-type' = 'BlockBlob'
        }
        $response = Invoke-RestMethod -Uri $uri -Method Put -Headers $headers -InFile $FileName -ContentType "application/json"
    }
    catch {
        Write-ErrorRunbook
        throw [CustomException]::new( $CmdletName, "$($response.StatusCode) StatusCode calling '$url'")
    }
}
#EndRegion ----------------------------------------------------- [Functions] ----------------------------------------------

try {
    $html = ConvertFrom-Html -URI 'https://learn.microsoft.com/en-us/officeupdates/update-history-microsoft365-apps-by-date'
    #$tableRows = ($html.SelectNodes('//table/tbody/tr'))
    $table = ($html.SelectNodes('//table'))
    $table1 = ($table[1].ChildNodes | Where-Object { $_.Name -eq 'tbody' })
    #$table1tr = $table1.ChildNodes | Where-Object { $_.Name -eq 'tr' }
    $TableData = New-Object System.Collections.Generic.List[System.Object]
    $table1.ChildNodes | Where-Object { $_.Name -eq 'tr' } | ForEach-Object {
        $currenttd = $_.ChildNodes | Where-Object { $_.name -eq 'td' }
        $OfficeCurrentPipeRelease = [PSCustomObject]@{
            Year                                       = $currenttd[0].InnerText 
            'Release Date'                             = $currenttd[1].InnerText 
            'Current Channel'                          = ($currenttd[2].ChildNodes | Where-Object { $_.name -eq 'a' }).InnerText
            'Monthly Enterprise Channel'               = ($currenttd[3].ChildNodes | Where-Object { $_.name -eq 'a' }).InnerText
            'Semi-Annual Enterprise Channel (Preview)' = ($currenttd[4].ChildNodes | Where-Object { $_.name -eq 'a' }).InnerText
            'Semi-Annual Enterprise Channel'           = ($currenttd[5].ChildNodes | Where-Object { $_.name -eq 'a' }).InnerText
        }
        $TableData.Add($OfficeCurrentPipeRelease)
    }
    if ($TableData) { $TableData | ConvertTo-Json | Out-File OfficeReleaseHistory.json } else { throw [CustomException]::new( "NoDataFromURL", "Parsing the webpage seems to fail... No Release history table data found.") }
    $JsonFile = Get-Item .\OfficeReleaseHistory.json -ErrorAction Stop
  #  Add-FileToBlobStorage -FilePath OfficeReleaseHistory.json
    Write-Output "Added OfficeReleaseHistory.json on blob storage. File size: $($JsonFile.Length) bytes, LastWriteTimeUtc: $($JsonFile.LastWriteTimeUtc)"
    
    $VerToBuild = New-Object -TypeName PSObject
    foreach ($ReleaseDateObj in $TableData) {
        $ReleaseDateObj.psobject.Properties | ForEach-Object {
            if ($_.Value -like '*Version*') {
                $_.value | ForEach-Object {
                    $Ver = $_ | Select-String -Pattern 'Version (\d+)' | ForEach-Object { $_.Matches.Groups[1].Value }
                    $Build = $_ | Select-String -Pattern '\(Build (.*)\)' | ForEach-Object { $_.Matches.Groups[1].Value }
                    if (-not [string]::IsNullOrWhitespace($Ver)) {$Ver = $Ver.Trim()}
                    if (-not [string]::IsNullOrWhitespace($Build)) {$Build = $Build.Trim()}
                    if ([bool]($VerToBuild.PSobject.Properties.name -eq $Ver)) {
                        if ($VerToBuild.$Ver -notcontains $Build) {
                            $VerToBuild.$Ver += $Build
                        }
                    } else {
                        $VerToBuild | Add-Member -MemberType NoteProperty -Name $Ver -Value @($Build)
                    }
                }
            }
        }
    }
    if ([string]::IsNullOrWhitespace($VerToBuild.PSObject.Properties)) {
        throw [CustomException]::new("EmptyVerToBuild", "The VerToBuild object is empty. No version and build information found.")
    }
    else { $VerToBuild | ConvertTo-Json | Out-File OfficeVersionToBuildMatch.json }
    $JsonFile = Get-Item .\OfficeVersionToBuildMatch.json -ErrorAction Stop
    Add-FileToBlobStorage -FilePath OfficeVersionToBuildMatch.json
    Write-Output "Added OfficeVersionToBuildMatch.json on blob storage. File size: $($JsonFile.Length) bytes, LastWriteTimeUtc: $($JsonFile.LastWriteTimeUtc)"
}
catch {
    Write-ErrorRunbook
    throw [CustomException]::new( "Terminating error", "Something went wrong... check the verbose log")
}

