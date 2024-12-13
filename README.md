# MicrosoftAppsReleaseHistoryToJSON
Used to convert the release history of commonly used Microsoft Apps from HTML to JSON. The resulting JSON format simplifies usage in pipelines or Compliance Policies.
<br>The blob storage  upload functionality for the resulting JSON files is also included.
<br>Web scraping is performed with the help of HTMLAgilityPack NuGet Package.
<br>*As per 2025 Microsoft does not provide an API for the realese history of their applications.*

### Microsoft 365 Apps (formerly known as Office 365)
This tool extracts the "Latest Release" table from the following page and converts it into JSON:
https://learn.microsoft.com/en-us/officeupdates/update-history-microsoft365-apps-by-date#supported-versions
