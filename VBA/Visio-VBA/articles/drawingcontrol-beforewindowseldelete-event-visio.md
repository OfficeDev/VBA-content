---
title: DrawingControl.BeforeWindowSelDelete Event (Visio)
ms.prod: visio
api_name:
- Visio.DrawingControl.BeforeWindowSelDelete
ms.assetid: 698ced64-c647-bee3-bd3d-31dc7386b7af
ms.date: 06/08/2017
---


# DrawingControl.BeforeWindowSelDelete Event (Visio)

Occurs before the shapes in the selection of a window are deleted.


## Syntax

Private Sub  _expression_ _**BeforeWindowSelDelete**( **_ByVal Window As [IVWINDOW]_** )

 _expression_ A variable that represents a **DrawingControl** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Window_|Required| **[IVWINDOW]**|The window that contains the selection that is going to be deleted.|

## Remarks

The  **BeforeWindowSelDelete** event fires if user interactions cause shapes in a window to be deleted. It doesn't fire if a program deletes shapes in a window by using the **Cut** method, for example.

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


