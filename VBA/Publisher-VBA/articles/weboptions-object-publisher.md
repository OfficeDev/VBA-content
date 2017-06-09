---
title: WebOptions Object (Publisher)
keywords: vbapb10.chm8323071
f1_keywords:
- vbapb10.chm8323071
ms.prod: publisher
api_name:
- Publisher.WebOptions
ms.assetid: 15358c46-f7ca-bc37-d7ef-7d4dbfee09a4
ms.date: 06/08/2017
---


# WebOptions Object (Publisher)

Represents the properties of a Web publication, including options for saving and encoding the publication, and enabling Web-safe fonts and font schemes. The  **WebOptions** object is a member of the **[Application](application-object-publisher.md)** object.
 


## Remarks

The properties of the  **WebOptions** object are used to specify the behavior of Web publications. This means that when any of these properties are modified, newly created Web publications inherit the modified properties.
 

 
Note that the  **WebOptions** object is available from print publications and Web publications. However, the properties of this object have no effect on print publications.
 

 

## Example

Use the  **[WebOptions](application-weboptions-property-publisher.md)** property on the **Application** object to return a **WebOptions** object. The following example sets an object variable equal to the Microsoft Publisher **WebOptions** object.
 

 

```
Dim theWO As WebOptions 
 
Set theWO = Application.WebOptions
```


## Properties



|**Name**|
|:-----|
|[AlwaysSaveInDefaultEncoding](weboptions-alwayssaveindefaultencoding-property-publisher.md)|
|[Application](weboptions-application-property-publisher.md)|
|[EmailAsImg](weboptions-emailasimg-property-publisher.md)|
|[EnableIncrementalUpload](weboptions-enableincrementalupload-property-publisher.md)|
|[Encoding](weboptions-encoding-property-publisher.md)|
|[OrganizeInFolder](weboptions-organizeinfolder-property-publisher.md)|
|[Parent](weboptions-parent-property-publisher.md)|
|[RelyOnVML](weboptions-relyonvml-property-publisher.md)|
|[ShowOnlyWebFonts](weboptions-showonlywebfonts-property-publisher.md)|

