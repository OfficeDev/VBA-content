---
title: Selection.LineStyle Property (Visio)
keywords: vis_sdr.chm11113845
f1_keywords:
- vis_sdr.chm11113845
ms.prod: visio
api_name:
- Visio.Selection.LineStyle
ms.assetid: 8bfba446-5987-58d1-54e2-5e861d7ce48d
ms.date: 06/08/2017
---


# Selection.LineStyle Property (Visio)

Specifies the line style for an object. Read/write.


## Syntax

 _expression_ . **LineStyle**

 _expression_ A variable that represents a **Selection** object.


### Return Value

String


## Remarks

Setting a style to a nonexistent style generates an error. Setting one kind of style to an existing style of another kind (for example, setting the  **LineStyle** property to a fill style) does nothing. Setting one kind of style to an existing style that has more than one set of attributes changes only the attributes for that component. For example, setting the **LineStyle** property to a style that has line, text, and fill attributes changes only the line attributes.

To preserve a shape's local formatting, use the  **LineStyleKeepFmt** property.

Beginning with Microsoft Visio 2002, setting  **LineStyle** to a zero-length string ("") will cause the master's style to be reapplied to the selection or shape. (Earlier versions generate a "no such style" exception.) If the selection or shape has no master, its style remains unchanged.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **LineStyle** property to set the line style for a particular shape.


```vb
 
Public Sub LineStyle_Example() 
 
 Dim vsoShape As Visio.Shape 
 
 Set vsoShape = ActivePage.DrawLine(5, 4, 7.5, 1) 
 vsoShape.LineStyle = "Guide" 
 
End Sub
```


