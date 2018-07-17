---
title: InvisibleApp.QueryCancelStyleDelete Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.QueryCancelStyleDelete
ms.assetid: 89993c25-5e0c-0dac-4e90-2dd08d4e7360
ms.date: 06/08/2017
---


# InvisibleApp.QueryCancelStyleDelete Event (Visio)

Occurs before the application deletes a style in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.


## Syntax

Private Sub  _expression_ _**QueryCancelStyleDelete**( **_ByVal Style As [IVSTYLE]_** )

 _expression_ A variable that represents an **InvisibleApp** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _style_|Required| **[IVSTYLE]**|The style that is going to be deleted.|

## Remarks

A Microsoft Visio instance fires  **QueryCancelStyleDelete** after the user has directed the instance to delete a style.




- If any event handler returns  **True** (cancel), the instance fires **StyleDeleteCanceled** and does not delete the style.
    
- If all handlers return  **False** (don't cancel), the instance fires **BeforeStyleDelete** and then deletes the style.
    


While a Visio instance is firing a query or cancel event, it responds to inquiries from client code but refuses to perform operations. Client code can show forms or message boxes while responding to a query or cancel event.

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


