---
title: LineFormat.DashStyle Property (Publisher)
keywords: vbapb10.chm3408132
f1_keywords:
- vbapb10.chm3408132
ms.prod: publisher
api_name:
- Publisher.LineFormat.DashStyle
ms.assetid: c2904350-89c1-2fc0-5bae-86f5193c8732
ms.date: 06/08/2017
---


# LineFormat.DashStyle Property (Publisher)

Returns or sets an  **MsoLineDashStyle** constant indicating the dash style for the specified line. Read/write.


## Syntax

 _expression_. **DashStyle**

 _expression_A variable that represents a  **LineFormat** object.


### Return Value

MsoLineDashStyle


## Remarks

The  **DashStyle** property value can be one of the ** [MsoLineDashStyle](http://msdn.microsoft.com/library/aba7f9d7-1689-c4a8-3b1e-e8dfb4a81d44%28Office.15%29.aspx)** constants declared in the Microsoft Office type library.


## Example

This example adds a blue dashed line to the active publication.


```vb
With ActiveDocument.Pages(1).Shapes _ 
 .AddLine(BeginX:=10, BeginY:=10, _ 
 EndX:=250, EndY:=250).Line 
 .DashStyle = msoLineDashDotDot 
 .ForeColor.RGB = RGB(50, 0, 128) 
End With 

```


