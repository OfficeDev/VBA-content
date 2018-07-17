---
title: Shape.BeforeShapeTextEdit Event (Visio)
keywords: vis_sdr.chm11219380
f1_keywords:
- vis_sdr.chm11219380
ms.prod: visio
api_name:
- Visio.Shape.BeforeShapeTextEdit
ms.assetid: f64b57b6-c92c-dd17-9698-211d9ca2fe83
ms.date: 06/08/2017
---


# Shape.BeforeShapeTextEdit Event (Visio)

Occurs before a shape is opened for text editing in the user interface.


## Syntax

Private Sub  _expression_ _**BeforeShapeTextEdit**( **_ByVal Shape As [IVSHAPE]_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Shape_|Required| **[IVSHAPE]**|The shape that is going to be opened for text editing.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


