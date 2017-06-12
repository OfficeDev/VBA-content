---
title: Document.GroupCanceled Event (Visio)
keywords: vis_sdr.chm10562005
f1_keywords:
- vis_sdr.chm10562005
ms.prod: visio
api_name:
- Visio.Document.GroupCanceled
ms.assetid: 135d8176-2c26-12aa-5bff-0df205e0640f
ms.date: 06/08/2017
---


# Document.GroupCanceled Event (Visio)

Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelGroup** event.


## Syntax

Private Sub  _expression_ _**GroupCanceled**( **_ByVal Selection As [IVSELECTION]_** )

 _expression_ A variable that represents a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Selection_|Required| **[IVSELECTION]**|The selection of shapes that was going to be grouped.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


