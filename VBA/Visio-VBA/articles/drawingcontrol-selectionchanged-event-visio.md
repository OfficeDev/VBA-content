---
title: DrawingControl.SelectionChanged Event (Visio)
ms.prod: visio
api_name:
- Visio.DrawingControl.SelectionChanged
ms.assetid: c655fa19-1e2b-8602-455e-34f6252386d2
ms.date: 06/08/2017
---


# DrawingControl.SelectionChanged Event (Visio)

Occurs after a set of shapes selected in a window changes.


## Syntax

Private Sub  _expression_ _**SelectionChanged**( **_ByVal Window As [IVWINDOW]_** )

 _expression_ A variable that represents a **DrawingControl** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window in which the selection changed.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


