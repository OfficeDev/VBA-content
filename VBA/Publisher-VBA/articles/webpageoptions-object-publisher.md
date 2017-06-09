---
title: WebPageOptions Object (Publisher)
keywords: vbapb10.chm548863
f1_keywords:
- vbapb10.chm548863
ms.prod: publisher
api_name:
- Publisher.WebPageOptions
ms.assetid: 694b56ce-1c2d-8202-25b7-19e55aadb0fd
ms.date: 06/08/2017
---


# WebPageOptions Object (Publisher)

Represents the properties of a single Web page within a Web publication, including options for adding the title and description of the page, background sounds, in addition to other options. The  **WebPageOptions** object is a member of the **[Page](page-object-publisher.md)** object.
 


## Remarks

Note that the  **WebPageOptions** object is only available when the active publication is a Web publication. A run-time error is returned if trying to access this object from a print publication.
 

 

## Example

Use the  **[WebPageOptions](page-webpageoptions-property-publisher.md)** property on the **Page** object to return a **WebPageOptions** object. Use the **[Description](webpageoptions-description-property-publisher.md)** property to set the description of a specified Web page. The following example sets the description for the second page of the active Web publication.
 

 

```
Dim theWPO As WebPageOptions 
 
Set theWPO = ActiveDocument.Pages(2).WebPageOptions 
 
With theWPO 
 .Description = "Company Profile" 
End With
```


## Methods



|**Name**|
|:-----|
|[SetBackgroundSoundRepeat](webpageoptions-setbackgroundsoundrepeat-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](webpageoptions-application-property-publisher.md)|
|[BackgroundSound](webpageoptions-backgroundsound-property-publisher.md)|
|[BackgroundSoundLoopCount](webpageoptions-backgroundsoundloopcount-property-publisher.md)|
|[BackgroundSoundLoopForever](webpageoptions-backgroundsoundloopforever-property-publisher.md)|
|[Description](webpageoptions-description-property-publisher.md)|
|[IncludePageOnNewWebNavigationBars](webpageoptions-includepageonnewwebnavigationbars-property-publisher.md)|
|[Keywords](webpageoptions-keywords-property-publisher.md)|
|[Parent](webpageoptions-parent-property-publisher.md)|
|[PublishFileName](webpageoptions-publishfilename-property-publisher.md)|

