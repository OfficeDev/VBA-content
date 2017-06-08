---
title: Documents.QueryCancelMasterDelete Event (Visio)
keywords: vis_sdr.chm10619310
f1_keywords:
- vis_sdr.chm10619310
ms.prod: visio
api_name:
- Visio.Documents.QueryCancelMasterDelete
ms.assetid: 44720e86-63cf-c392-e981-672969e09f25
ms.date: 06/08/2017
---


# Documents.QueryCancelMasterDelete Event (Visio)

Occurs before the application deletes a master in response to a user action in the interface. If any event handler returns  **True** , the operation is canceled.


## Syntax

Private Sub  _expression_ _**QueryCancelMasterDelete**( **_ByVal Master As [IVMASTER]_** )

 _expression_ A variable that represents a **Documents** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Master_|Required| **[IVMASTER]**|The master that is going to be deleted.|

## Remarks

A Microsoft Visio instance fires  **QueryCancelMasterDelete** after the user has directed the instance to delete a master.




- If any event handler returns  **True** (cancel), the instance fires **MasterDeleteCanceled** and does not delete the master.
    
- If all handlers return  **False** (don't cancel), the instance fires **BeforeMasterDelete** and then deletes the master.
    


While a Visio instance is firing a query or cancel event, it responds to inquiries from client code but refuses to perform operations. Client code can show forms or message boxes while responding to a query or cancel event.

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


