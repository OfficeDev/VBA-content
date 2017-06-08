---
title: InvisibleApp.QueryCancelSuspend Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.QueryCancelSuspend
ms.assetid: 49e6dbe2-f1d9-5743-11d2-c64e1d98475d
ms.date: 06/08/2017
---


# InvisibleApp.QueryCancelSuspend Event (Visio)

Occurs before the operating system enters a suspended state. If any event handler returns  **True** , the Microsoft Visio instance will deny the operating system's request.


## Syntax

Private Sub  _expression_ _**QueryCancelSuspend**( **_ByVal app As [IVAPPLICATION]_** )

 _expression_ A variable that represents an **InvisibleApp** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Visio that responds to the operating system request.|

## Remarks

 You will typically respond **False** and allow the operating system to enter a suspended state. If you have open network files, you can close them when you receive the **BeforeSuspend** event. If you have open network files that you cannot close, you can return **True** and Visio will deny the operating system's request.




- If any event handler returns  **True** (cancel), the instance fires **SuspendCanceled** and does not enter a suspended state.
    
- If all handlers return  **False** (don't cancel), the instance fires **BeforeSuspend** and then enters a suspended state.
    


If your solution runs outside the Visio process, you cannot be assured of receiving this event. For this reason, you should monitor window messages in your program.

While a Visio instance is firing a query or cancel event, it responds to inquiries from client code but refuses to perform operations. Client code can show forms or message boxes while responding to a query or cancel event.

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


## Example

This VBA macro shows how to capture the  **QueryCancelSuspend** event and allow the operating system to suspend. Declare a **WithEvents** variable to capture events fired by the **Application** object.


```vb
 
Public WithEvents vsoApplication As Visio.Application  
  
Private Function vsoApplication_QueryCancelSuspend(ByVal _ 
    IVisioApplication As IVApplication) As Boolean 
  
    'You agree to let the operating system suspend.  
    vsoApplication_QueryCancelSuspend = False 
  
End Function
```


