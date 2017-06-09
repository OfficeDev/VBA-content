---
title: Documents.DocumentCloseCanceled Event (Visio)
keywords: vis_sdr.chm10619340
f1_keywords:
- vis_sdr.chm10619340
ms.prod: visio
api_name:
- Visio.Documents.DocumentCloseCanceled
ms.assetid: 03bc417c-2a80-551e-d307-1035dd6088d6
ms.date: 06/08/2017
---


# Documents.DocumentCloseCanceled Event (Visio)

Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelDocumentClose** event.


## Syntax

Private Sub  _expression_ _**DocumentCloseCanceled**( **_ByVal doc As [IVDOCUMENT]_** )

 _expression_ A variable that represents a **Documents** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _doc_|Required| **[IVDOCUMENT]**|The document that was going to be closed.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


