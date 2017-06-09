---
title: Document.BeforeDocumentClose Event (Visio)
keywords: vis_sdr.chm10519025
f1_keywords:
- vis_sdr.chm10519025
ms.prod: visio
api_name:
- Visio.Document.BeforeDocumentClose
ms.assetid: e35f9593-f5ee-f84b-95e6-f23a899c0d6d
ms.date: 06/08/2017
---


# Document.BeforeDocumentClose Event (Visio)

Occurs before a document is closed.


## Syntax

Private Sub  _expression_ _**BeforeDocumentClose**( **_ByVal doc As [IVDOCUMENT]_** )

 _expression_ A variable that represents a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _doc_|Required| **[IVDOCUMENT]**|The document that is going to be closed.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


