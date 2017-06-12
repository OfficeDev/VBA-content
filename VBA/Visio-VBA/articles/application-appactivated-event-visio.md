---
title: Application.AppActivated Event (Visio)
ms.prod: visio
api_name:
- Visio.Application.AppActivated
ms.assetid: 150864ab-574a-6556-a56a-8ca619796062
ms.date: 06/08/2017
---


# Application.AppActivated Event (Visio)

Occurs after a Microsoft Visio instance becomes active.


## Syntax

Private Sub  _expression_ _**AppActivated**( **_ByVal app As [IVAPPLICATION]_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _app_|Required| **[IVAPPLICATION]**|The instance of Visio that becomes the active application.|

## Remarks

The  **AppActivated** event indicates that an instance of Visio has become the active application on the Microsoft Windows desktop. The **AppActivated** event is different from the **AppObjectActivated** event, which occurs after an instance of Visio becomes active?the instance of Visio that is retrieved by the **GetObject** method in a Microsoft Visual Basic program.

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


