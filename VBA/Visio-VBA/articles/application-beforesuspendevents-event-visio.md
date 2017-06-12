---
title: Application.BeforeSuspendEvents Event (Visio)
ms.prod: visio
api_name:
- Visio.Application.BeforeSuspendEvents
ms.assetid: a6879424-40d8-e517-aad0-f31aa84a49f6
ms.date: 06/08/2017
---


# Application.BeforeSuspendEvents Event (Visio)

Occurs before firing of events is suspended.


## Syntax

Private Sub  _expression_ _**BeforeSuspendEvents**( **_ByVal app As [IVAPPLICATION]_** )

 _expression_ An expression that returns a **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Microsoft Visio in which firing of events is going to be suspended.|

### Return Value

nothing


## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


