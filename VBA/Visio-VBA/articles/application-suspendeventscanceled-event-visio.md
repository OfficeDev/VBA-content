---
title: Application.SuspendEventsCanceled Event (Visio)
ms.prod: visio
api_name:
- Visio.Application.SuspendEventsCanceled
ms.assetid: 33892ba1-90b2-30ee-d355-e3c353749ea8
ms.date: 06/08/2017
---


# Application.SuspendEventsCanceled Event (Visio)

Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelSuspendEvents** event. .


## Syntax

Private Sub  _expression_ _**SuspendEventsCanceled**( **_ByVal app As_** )

 _expression_ An expression that returns a **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Microsoft Visio in which firing of events was going to be suspended.|

### Return Value

nothing


## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


