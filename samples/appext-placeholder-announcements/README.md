# Announcements SharePoint Framework application customizer

## Summary

Sample SharePoint Framework application customizer showing organizational announcements.

## Used SharePoint Framework Version

![SPFx v1.3.0](https://img.shields.io/badge/SPFx-1.3.0-green.svg)

## Applies to

* [SharePoint Framework Extensions ](https://dev.office.com/sharepoint/docs/spfx/extensions/overview-extensions)
* [Office 365 developer tenant](http://dev.office.com/sharepoint/docs/spfx/set-up-your-developer-tenant)

## Solution

Solution|Author(s)
--------|---------
react-announcements|Brian Caauwe (MCSM, [Avtex](https://avtex.com), @bcaauwe)

## Version history

Version|Date|Comments
-------|----|--------
1.0.0|November 17, 2017|Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Prerequisites

* Office 365 tenant with a modern site collection and a list of announcements

## Minimal Path to Awesome

* clone this repo
* in the command line run
  * `npm i`
  * `gulp serve --nobrowser`
* add parameters to specify the URL of the site and the name of the list where announcements are stored (see example below)
* in the web browser
  * navigate to the modern site
  * to the URL of the site add the previously copied debug query string parameters

## Features

This project contains sample SharePoint Framework application customizer built using React and Office UI Fabric React. The application customizer retrieves organizational announcements and displays them to the user until she acknowledges that she has seen them. Which announcements have been seen is stored in the local browser storage.

This sample illustrates the following concepts on top of the SharePoint Framework:

* using Office UI Fabric React to build SharePoint Framework application customizers that seamlessly integrate with SharePoint
* using React to build SharePoint Framework application customizers
* retrieving information from SharePoint using the SPHttpClient
* rendering information in page placeholders
* storing information in the local browser storage
* passing configuration parameters into application customizers

### Available configuration parameters

Parameter | Type | Possible values | Description
----------|------|-----------------|------------
`siteUrl`|string|absolute or server-relative URL|URL of the site where the list with announcements is located
`listName`|string|any string|Name of the list where the announcements are stored

### Announcements list structure

Column name|Type|Description
-----------|----|-----------
`ID`|Number|Auto-generated item ID, used to track which announcements have been seen by the user
`Title`|Single-line of text|Announcement title, displayed in bold
`Announcement`|Multiple lines of text (plain-text)|Announcement body. All absolute URLs are automatically turned into an HTML hyperlink
`Category`|Choice|Specifies if the category of the announcement. Possible values are: Blocked, Error, Info, Remove, Success or Warning

Show organizational announcements stored in the **Announcements** list in the root site collection:

```text
?loadSPFX=true&debugManifestsFile=https://localhost:4321/temp/manifests.js&customActions={"01f1bfa7-1224-4ea9-8db1-438f8fe01511":{"location":"ClientSideExtension.ApplicationCustomizer","properties":{"siteUrl":"https://contoso.sharepoint.com","listName":"Announcements"}}}
```