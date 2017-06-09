---
title: InvisibleApp.QueryCancelPageDelete Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.QueryCancelPageDelete
ms.assetid: d3dd09f7-b3b8-cc99-c060-09a9c06c36a3
ms.date: 06/08/2017
---


# InvisibleApp.QueryCancelPageDelete Event (Visio)

Occurs before the application deletes a page in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.


## Syntax

Private Sub  _expression_ _**QueryCancelPageDelete**( **_ByVal Page As [IVPAGE]_** )

 _expression_ A variable that represents an **InvisibleApp** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Page_|Required| **[IVPAGE]**|The page that is going to be deleted.|

## Remarks

A Visio instance fires  **QueryCancelPageDelete** after the user has directed the instance to delete a page.




- If any event handler returns  **True** (cancel), the instance fires **PageDeleteCanceled** and does not delete the page.
    
- If all handlers return  **False** (don't cancel) the instance fires **BeforePageDelete** and then deletes the page.
    


While a Microsoft Visio instance is firing a query or cancel event, it responds to inquiries from client code but refuses to perform operations. Client code can show forms or message boxes while responding to a query or cancel event.

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


