---
title: Documents.StyleAdded Event (Visio)
keywords: vis_sdr.chm10619245
f1_keywords:
- vis_sdr.chm10619245
ms.prod: visio
api_name:
- Visio.Documents.StyleAdded
ms.assetid: e2ba6aca-f07c-8c0e-20fd-d4ad1b1c8c57
ms.date: 06/08/2017
---


# Documents.StyleAdded Event (Visio)

Occurs after a new style is added to a document.


## Syntax

Private Sub  _expression_ _**StyleAdded**( **_ByVal Style As [IVSTYLE]_** )

 _expression_ A variable that represents a **Documents** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Style_|Required| **[IVSTYLE]**|The style that was added to the document.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


