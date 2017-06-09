---
title: Master.ShapeExitedTextEdit Event (Visio)
keywords: vis_sdr.chm10719385
f1_keywords:
- vis_sdr.chm10719385
ms.prod: visio
api_name:
- Visio.Master.ShapeExitedTextEdit
ms.assetid: 401f6d32-d1fb-f019-52a3-d553b8516ecf
ms.date: 06/08/2017
---


# Master.ShapeExitedTextEdit Event (Visio)

Occurs after a shape is no longer open for interactive text editing.


## Syntax

Private Sub  _expression_ _**ShapeExitedTextEdit**( **_ByVal Shape As [IVSHAPE]_** )

 _expression_ A variable that represents a **Master** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Shape_|Required| **[IVSHAPE]**|The shape that was closed for text editing.|

## Remarks

If you're using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](http://msdn.microsoft.com/library/de8f5c7a-421d-ebcf-22b6-4310a202ef64%28Office.15%29.aspx).


