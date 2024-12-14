# MicrosoftAppsReleaseHistoryToJSON
Used to convert the release history of commonly used Microsoft Apps from HTML to JSON. The resulting JSON format simplifies usage in pipelines or Compliance Policies.
<br>The blob storage  upload functionality for the resulting JSON files is also included.
<br>Web scraping is performed with the help of HTMLAgilityPack NuGet Package.
<br>*As per 2025 Microsoft does not provide an API for the realese history of their applications.*

### Microsoft 365 Apps (formerly known as Office 365)
* M365AppsLatestRelease-ToJson.ps1 extracts the "Latest Release" table from the following page and converts it into JSON:
https://learn.microsoft.com/en-us/officeupdates/update-history-microsoft365-apps-by-date#supported-versions
<br>* M365AppsLatestReleaseHistory-ToJson.ps1 extracts the "Version History" table from the following page and converts it into JSON:
https://learn.microsoft.com/en-us/officeupdates/update-history-microsoft365-apps-by-date#version-history
<br> Since O365 does not write the App version on the machine that is installed on, this script will also utilize the Version History table to create a JSON for matching versions to builds
<br> Both files are to determine the O365 version a machine is running on.
