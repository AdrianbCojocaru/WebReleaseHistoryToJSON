<#

.Version 1.0.0

.Author Adrian Cojocaru

.Date 28-Aug-2024

.Synopsis
    Convert the Microsoft Edge release schedule webpage to Json

.Description
    Convert https://learn.microsoft.com/en-us/deployedge/microsoft-edge-release-schedule to Json with the help of PowerHTML module v0.2.0
    The resulting Json is uploaded on blob storage so it can be read from a client device.
    The json file can be used further to determine if MS Edge is on the latest version.

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
    $html = ConvertFrom-Html -URI 'https://learn.microsoft.com/en-us/deployedge/microsoft-edge-release-schedule'
    $tableRows = ($html.SelectNodes('//table/tbody/tr'))
    $TableData = New-Object System.Collections.Generic.List[System.Object]
    $tableRows | ForEach-Object {
        $NodeElement = $_.InnerText -split '\r?\n'
        $VersionInt = $NodeElement[1]
        $regex = [regex] "(?=$VersionInt)"
        $BetaChannelRelease = $regex.Split($NodeElement[3])
        $StableChannelRelease = $regex.Split($NodeElement[4])
        $ExtendedStableChannelRelease = $regex.Split($NodeElement[5])
        $EdgeRelease = [PSCustomObject]@{
            Version                             = $NodeElement[1]
            ReleaseStatus                       = $NodeElement[2]
            BetaChannelReleaseWeek              = $BetaChannelRelease[0]
            BetaChannelReleaseVersion           = $BetaChannelRelease[1]
            StableChannelReleaseWeek            = $StableChannelRelease[0]
            StableChannelReleaseVersion         = $StableChannelRelease[1]
            ExtendedStableChannelReleaseWeek    = $ExtendedStableChannelRelease[0]
            ExtendedStableChannelReleaseVersion = $ExtendedStableChannelRelease[1]
        }
        $TableData.Add($EdgeRelease) 
    } 
    if ($TableData) { $TableData | ConvertTo-Json | Out-File EdgeReleases.json } else { throw [CustomException]::new( "NoDataFromURL", "Parsing the webpage seems to fail... No Release history table data found.") }
    $JsonFile = Get-Item .\EdgeReleases.json -ErrorAction Stop
    Add-FileToBlobStorage -FilePath EdgeReleases.json
    Write-Output "Added EdgeReleases.json on blob storage. File size: $($JsonFile.Length) bytes, LastWriteTimeUtc: $($JsonFile.LastWriteTimeUtc)"
}
catch {
    Write-ErrorRunbook
    throw [CustomException]::new( "Terminating error", "Something went wrong... check the verbose log")
}

