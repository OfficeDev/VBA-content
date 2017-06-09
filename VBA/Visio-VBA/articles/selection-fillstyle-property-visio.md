---
title: Selection.FillStyle Property (Visio)
keywords: vis_sdr.chm11113525
f1_keywords:
- vis_sdr.chm11113525
ms.prod: visio
api_name:
- Visio.Selection.FillStyle
ms.assetid: efdf51ba-7d0a-d5c0-5a39-d22d7a79a053
ms.date: 06/08/2017
---


# Selection.FillStyle Property (Visio)

Returns or sets the fill style for an object. Read/write.


## Syntax

 _expression_ . **FillStyle**

 _expression_ A variable that represents a **Selection** object.


### Return Value

String


## Remarks

Setting the  **FillStyle** property is equivalent to selecting a style from the **Style** list in the **Fill** dialog box (right-click the shape, point to **Format**, and then click  **Fill**).

Setting a style to a nonexistent style generates an error. Setting one type of style to another type (for example, setting the  **FillStyle** property to a line style) does nothing. Setting one type of style to another type that has more than one set of attributes changes only the appropriate attributes. For example, setting the **FillStyle** property to a style that has line, text, and fill attributes changes only the fill attributes.

To preserve a shape's local formatting, use the  **FillStyleKeepFmt** property.

Beginning with Microsoft Visio 2002, setting the  **FillStyle** property to a zero-length string ("") causes the master's style to be reapplied to the selection or shape. (Earlier versions generate a "no such style" exception.) If the selection or shape has no master, its style remains unchanged.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to draw an oval and set its  **FillStyle** property. To run this macro, open a drawing based on the **Basic Diagram** template.


```vb
 
Public Sub FillStyle_Example() 
 
 Dim vsoShape As Visio.Shape 
 
 Set vsoShape = ActivePage.DrawOval(1.5, 10.5, 7.5, 6.5) 
 vsoShape.FillStyle = "Basic" 
 
End Sub
```


