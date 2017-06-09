---
title: Application.StyleChanged Event (Visio)
ms.prod: visio
api_name:
- Visio.Application.StyleChanged
ms.assetid: f56680b3-71c3-91c6-23d0-7d5840f9aeb5
ms.date: 06/08/2017
---


# Application.StyleChanged Event (Visio)

Occurs after the name of a style is changed or a change to the style propagates to objects to which the style is applied.


## Syntax

Private Sub  _expression_ _**StyleChanged**( **_ByVal Style As [IVSTYLE]_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _style_|Required| **[IVSTYLE]**|The style that changed.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


