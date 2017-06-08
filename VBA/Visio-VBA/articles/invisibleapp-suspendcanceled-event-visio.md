---
title: InvisibleApp.SuspendCanceled Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.SuspendCanceled
ms.assetid: 5c266211-8686-85e8-f059-38e3cdab4211
ms.date: 06/08/2017
---


# InvisibleApp.SuspendCanceled Event (Visio)

Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelSuspend** event.


## Syntax

Private Sub  _expression_ _**SuspendCanceled**( **_ByVal app As [IVAPPLICATION]_** )

 _expression_ A variable that represents an **InvisibleApp** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Microsoft Visio that was going to be suspended.|

## Remarks

If your solution runs outside the Visio process, you cannot be assured of receiving this event. For this reason, you should monitor window messages in your program.

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


