---
title: Master.BeforeMasterDelete Event (Visio)
keywords: vis_sdr.chm10719040
f1_keywords:
- vis_sdr.chm10719040
ms.prod: visio
api_name:
- Visio.Master.BeforeMasterDelete
ms.assetid: 46b455db-9165-0ed4-ebf3-15e1794313be
ms.date: 06/08/2017
---


# Master.BeforeMasterDelete Event (Visio)

Occurs before a master is deleted from a document.


## Syntax

Private Sub  _expression_ _**BeforeMasterDelete**( **_ByVal Master As [IVMASTER]_** )

 _expression_ A variable that represents a **Master** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Master_|Required| **[IVMASTER]**|The master that is going to be deleted.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


