---
title: Document.MasterAdded Event (Visio)
keywords: vis_sdr.chm10519170
f1_keywords:
- vis_sdr.chm10519170
ms.prod: visio
api_name:
- Visio.Document.MasterAdded
ms.assetid: 5637df50-5174-03d4-a07f-cc7aeb92d0fa
ms.date: 06/08/2017
---


# Document.MasterAdded Event (Visio)

Occurs after a new master is added to a document.


## Syntax

Private Sub  _expression_ _**MasterAdded**( **_ByVal Master As [IVMASTER]_** )

 _expression_ A variable that represents a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Master_|Required| **[IVMASTER]**|The master that was added to the document.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


