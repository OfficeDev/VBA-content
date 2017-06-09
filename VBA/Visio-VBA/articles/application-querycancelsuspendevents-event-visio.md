---
title: Application.QueryCancelSuspendEvents Event (Visio)
ms.prod: visio
api_name:
- Visio.Application.QueryCancelSuspendEvents
ms.assetid: 886fa424-67b3-6a4d-f0bb-99ee646b0753
ms.date: 06/08/2017
---


# Application.QueryCancelSuspendEvents Event (Visio)

Occurs before the application suspends events in response to client code. If any event handler returns  **True** , the operation is canceled.


## Syntax

Private Sub  _expression_ _**QueryCancelSuspendEvents**( **_ByVal app As [IVAPPLICATION]_** )

 _expression_ An expression that returns a **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Microsoft Visio in which firing of events is going to be suspended.|

### Return Value

nothing


## Remarks

A Visio instance fires  **QueryCancelSuspendEvents** after client code has directed the instance to suspend events.




- If any event handler returns  **True** (cancel), the instance fires **SuspendEventsCanceled** and does not suspend events.
    
- If all handlers return  **False** (don't cancel), the suspension occurs.
    


While a Visio instance is firing a query or cancel event, it responds to inquiries from client code but refuses to perform operations. Client code can show forms or message boxes while responding to a query or cancel event.

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


