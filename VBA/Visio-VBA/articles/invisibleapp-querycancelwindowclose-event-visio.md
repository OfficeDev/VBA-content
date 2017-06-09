---
title: InvisibleApp.QueryCancelWindowClose Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.QueryCancelWindowClose
ms.assetid: 78fb3502-4827-5add-1ff1-5dd3e3b6f143
ms.date: 06/08/2017
---


# InvisibleApp.QueryCancelWindowClose Event (Visio)

Occurs before the application closes a window in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.


## Syntax

Private Sub  _expression_ _**QueryCancelWindowClose**( **_ByVal Window As [IVWINDOW]_** )

 _expression_ A variable that represents an **InvisibleApp** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window that is going to be closed.|

## Remarks

A Microsoft Visio instance fires  **QueryCancelWindowClose** after the user has directed the instance to close a window.




- If any event handler returns  **True** (cancel), the instance fires **WindowCloseCanceled** and does not close the window.
    
- If all handlers return  **False** (don't cancel), the instance fires **BeforeWindowClosed** and then closes the window.
    


While a Visio instance is firing a query or cancel event, it responds to inquiries from client code but it refuses to perform operations. Client code can show forms or message boxes while responding to a query or cancel event.

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


