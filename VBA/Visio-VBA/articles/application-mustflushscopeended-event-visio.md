---
title: Application.MustFlushScopeEnded Event (Visio)
ms.prod: visio
api_name:
- Visio.Application.MustFlushScopeEnded
ms.assetid: ba9ae16a-9cc6-79d6-d838-e5927937c142
ms.date: 06/08/2017
---


# Application.MustFlushScopeEnded Event (Visio)

Occurs after the Microsoft Visio instance is forced to flush its event queue.


## Syntax

Private Sub  _expression_ _**MustFlushScopeEnded**( **_ByVal app As [IVAPPLICATION]_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Visio that is forced to flush its event queue.|

## Remarks

This event, along with the  **MustFlushScopeBeginning** event, can be used to identify whether an event is being fired because Visio is forced to flush its event queue.

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


