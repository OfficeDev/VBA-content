---
title: WebNavigationBarHyperlinks.Add Method (Publisher)
keywords: vbapb10.chm8585220
f1_keywords:
- vbapb10.chm8585220
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarHyperlinks.Add
ms.assetid: 6cd0c43a-fec1-c9b8-dc86-00e1cc314087
ms.date: 06/08/2017
---


# WebNavigationBarHyperlinks.Add Method (Publisher)

Adds a new  **Hyperlink** object to the specified **WebNavigationBarHyperlinks** collection and returns the new **Hyperlink** object. .


## Syntax

 _expression_. **Add**( **_Address_**,  **_RelativePage_**,  **_PageID_**,  **_TextToDisplay_**,  **_Index_**)

 _expression_A variable that represents a  **WebNavigationBarHyperlinks** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Address|Optional| **String**|The address of the new hyperlink. If RelativePage is  **pbHlinkTargetTypeURL** (default) or **pbHlinkTargetTypeEmail**, Address must be specified or an error occurs.|
|RelativePage|Optional| **PbHlinkTargetType**|The type of hyperlink to add.|
|PageID|Optional| **Long**|The page ID of the destination page for the new hyperlink. If RelativePage is  **pbHlinkTargetTypePageID**, PageID must be specified or an error occurs. The page ID corresponds to the  [PageID](page-pageid-property-publisher.md) property of the destination page.|
|TextToDisplay|Optional| **String**|The display text of the new hyperlink. |
|Index|Optional| **Long**|The index of the new  **Hyperlink** object in the **WebNavigationBarHyperlinks** collection.|

### Return Value

Hyperlink


## Remarks

RelativePage can be one of these  [PbHlinkTargetType](pbhlinktargettype-enumeration-publisher.md) constants. The default is **pbHlinkTargetTypeURL**.



| **pbHlinkTargetTypeEmail**|
| **pbHlinkTargetTypeFirstPage**|
| **pbHlinkTargetTypeLastPage**|
| **pbHlinkTargetTypeNextPage**|
| **pbHlinkTargetTypePageID**|
| **pbHlinkTargetTypePreviousPage**|
| **pbHlinkTargetTypeURL**|

