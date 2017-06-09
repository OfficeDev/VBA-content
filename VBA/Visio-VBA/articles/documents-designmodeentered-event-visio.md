---
title: Documents.DesignModeEntered Event (Visio)
keywords: vis_sdr.chm10619110
f1_keywords:
- vis_sdr.chm10619110
ms.prod: visio
api_name:
- Visio.Documents.DesignModeEntered
ms.assetid: d3858366-1922-6462-498d-ba6d09219e7f
ms.date: 06/08/2017
---


# Documents.DesignModeEntered Event (Visio)

Occurs before a document enters design mode.


## Syntax

Private Sub  _expression_ _**DesignModeEntered**( **_ByVal doc As [IVDOCUMENT]_** )

 _expression_ A variable that represents a **Documents** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _doc_|Required| **[IVDOCUMENT]**|The document that is going to enter design mode.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


