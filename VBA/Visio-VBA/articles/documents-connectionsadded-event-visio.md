---
title: Documents.ConnectionsAdded Event (Visio)
keywords: vis_sdr.chm10619095
f1_keywords:
- vis_sdr.chm10619095
ms.prod: visio
api_name:
- Visio.Documents.ConnectionsAdded
ms.assetid: db864aa5-527d-9007-364f-f90c944ddda6
ms.date: 06/08/2017
---


# Documents.ConnectionsAdded Event (Visio)

Occurs after connections have been established between shapes.


## Syntax

Private Sub  _expression_ _**ConnectionsAdded**( **_ByVal Connects As [IVCONNECTS]_** )

 _expression_ A variable that represents a **Documents** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Connects_|Required| **[IVCONNECTS]**|The connections that were established.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).




 **Note**   You can use VBA **WithEvents** variables to sink the **ConnectionsDeleted** event.

For performance considerations, the  **Document** object's event set does not include the **ConnectionsAdded** event. To sink the **ConnectionsAdded** event from a **Document** object (and the **ThisDocument** object in a VBA project), you must use the **AddAdvise** method.


