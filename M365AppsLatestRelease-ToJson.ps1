<#

.Version 1.0.0

.Author Adrian

.Date 10-Dec-2024

.Synopsis
    Convert the Microsoft 365 Apps latest release HTML table to Json

.Description
    Convert https://learn.microsoft.com/en-us/officeupdates/update-history-microsoft365-apps-by-date#supported-versions (latest release) to Json with the help of the HTMLAgilityPack module
    The resulting Json can be uploaded on a blob storage so it can be read from the client.
    The json file is used to determine if O365 is on the latest version.

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
    $table1 = $html.SelectSingleNode('//table[1]')
    $tbody = $table1.ChildNodes | Where-Object { $_.name -eq 'tbody' }
    $TableData = New-Object System.Collections.Generic.List[System.Object]
    $tbody.ChildNodes | Where-Object { $_.Name -eq 'tr' } | ForEach-Object {
        $currenttd = $_.ChildNodes | Where-Object { $_.name -eq 'td' }
        $OfficeLatestRelease = [PSCustomObject]@{
            Channel                 = $currenttd[0].InnerText 
            Version                 = $currenttd[1].InnerText 
            Build                   = $currenttd[2].InnerText 
            LatestReleaseDate       = $currenttd[3].InnerText 
            VersionAvailabilityDate = $currenttd[4].InnerText 
            EndOfService            = $currenttd[5].InnerText
        }
        $TableData.Add($OfficeLatestRelease)
    }

  
    if ($TableData) { $TableData | ConvertTo-Json | Out-File OfficeLatestRelease.json } else { throw [CustomException]::new( "NoDataFromURL", "Parsing the webpage seems to fail... No Release history table data found.") }
    $JsonFile = Get-Item .\OfficeLatestRelease.json -ErrorAction Stop
    Add-FileToBlobStorage -FilePath OfficeLatestRelease.json
    Write-Output "Added OfficeLatestRelease.json on blob storage. File size: $($JsonFile.Length) bytes, LastWriteTimeUtc: $($JsonFile.LastWriteTimeUtc)"
}
catch {
    Write-ErrorRunbook
    throw [CustomException]::new( "Terminating error", "Something went wrong... check the verbose log")
}

