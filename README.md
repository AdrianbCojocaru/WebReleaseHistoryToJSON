# MicrosoftAppsReleaseHistoryToJSON
Used to convert the release history of commonly used Microsoft Apps from HTML to JSON. The resulting JSON format simplifies usage in pipelines or Compliance Policies.
<br>The blob storage  upload functionality for the resulting JSON files is also included.
<br>Web scraping is performed with the help of HTMLAgilityPack NuGet Package.
<br>*As per 2025 Microsoft does not provide an API for the realese history of their applications.*

### Microsoft 365 Apps (formerly known as Office 365)
* M365AppsLatestRelease-ToJson.ps1 extracts the "Latest Release" table from the following page and converts it into JSON:
https://learn.microsoft.com/en-us/officeupdates/update-history-microsoft365-apps-by-date#supported-versions
* M365AppsReleaseHistory-ToJson.ps1 extracts the "Version History" table from the following page and converts it into JSON:
https://learn.microsoft.com/en-us/officeupdates/update-history-microsoft365-apps-by-date#version-history
<br> O365 does not write any record of the App version on a machine where it is installed, only the build version is stamped. This script uses the Version History table to generate a second JSON file that maps app versions to build numbers.
<br> Both files are necesary to determine the O365 version that is installed a machine.

### Microsoft Edge
MicrosoftEdgeReleaseHistory-ToJson.ps1 used to convert the following "Microsoft Edge releases" table to JSON: https://learn.microsoft.com/en-us/deployedge/microsoft-edge-release-schedule#microsoft-edge-releases
<br>