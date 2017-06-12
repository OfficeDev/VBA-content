---
title: InvisibleApp.ShapeExitedTextEdit Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.ShapeExitedTextEdit
ms.assetid: 54e52c06-b7ab-f6c3-9c0d-6ee05da0e1f3
ms.date: 06/08/2017
---


# InvisibleApp.ShapeExitedTextEdit Event (Visio)

Occurs after a shape is no longer open for interactive text editing.


## Syntax

Private Sub  _expression_ _**ShapeExitedTextEdit**( **_ByVal Shape As [IVSHAPE]_** )

 _expression_ A variable that represents an **InvisibleApp** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Shape_|Required| **[IVSHAPE]**|The shape that was closed for text editing.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


