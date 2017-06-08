---
title: Documents.MasterDeleteCanceled Event (Visio)
keywords: vis_sdr.chm10619355
f1_keywords:
- vis_sdr.chm10619355
ms.prod: visio
api_name:
- Visio.Documents.MasterDeleteCanceled
ms.assetid: 4a32dd7f-a7c6-05d4-9be2-eed5900e339c
ms.date: 06/08/2017
---


# Documents.MasterDeleteCanceled Event (Visio)

Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelMasterDelete** event.


## Syntax

Private Sub  _expression_ _**MasterDeleteCanceled**( **_ByVal Master As [IVMASTER]_** )

 _expression_ A variable that represents a **Documents** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Master_|Required| **[IVMASTER]**|The master that was going to be deleted.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


