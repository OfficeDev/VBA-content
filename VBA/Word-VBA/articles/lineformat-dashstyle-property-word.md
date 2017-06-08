---
title: LineFormat.DashStyle Property (Word)
keywords: vbawd10.chm164233320
f1_keywords:
- vbawd10.chm164233320
ms.prod: word
api_name:
- Word.LineFormat.DashStyle
ms.assetid: 1dd61d77-d7fc-cb8d-5d44-38aca7073a68
ms.date: 06/08/2017
---


# LineFormat.DashStyle Property (Word)

Returns or sets the dash style for the specified line. Read/write  **MsoLineDashStyle** .


## Syntax

 _expression_ . **DashStyle**

 _expression_ Required. A variable that represents a **[LineFormat](lineformat-object-word.md)** object.


## Example

This example adds a blue dashed line to the active document.


```vb
Dim docActive As Document 
 
Set docActive = ActiveDocument 
 
With docActive.Shapes.AddLine(10, 10, 250, 250).Line 
 .DashStyle = msoLineDashDotDot 
 .ForeColor.RGB = RGB(50, 0, 128) 
End With
```


## See also


#### Concepts


[LineFormat Object](lineformat-object-word.md)

