---
title: Shapes.SelectAll Method (Word)
keywords: vbawd10.chm161415190
f1_keywords:
- vbawd10.chm161415190
ms.prod: word
api_name:
- Word.Shapes.SelectAll
ms.assetid: 2d907cfd-75ad-c29f-8ef8-85f810915ba8
ms.date: 06/08/2017
---


# Shapes.SelectAll Method (Word)

Selects all the shapes in a collection of shapes.


## Syntax

 _expression_ . **SelectAll**

 _expression_ Required. A variable that represents a **[Shapes](shapes-object-word.md)** collection.


## Remarks

This method does not select  **InlineShape** objects. You cannot use this method to select more than one canvas.


## Example

This example selects all the shapes in the active document.


```vb
Sub SelectAllShapes() 
 ActiveDocument.Shapes.SelectAll 
End Sub
```

This example selects all the shapes in the headers and footers of the active document and adds a red shadow to each shape.




```vb
Sub SelectAllHeaderShapes() 
 With ActiveDocument.ActiveWindow 
 .View.Type = wdPrintView 
 .ActivePane.View.SeekView = wdSeekCurrentPageHeader 
 End With 
 
 ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Shapes.SelectAll 
 
 With Selection.ShapeRange.Shadow 
 .Type = msoShadow10 
 .ForeColor.RGB = RGB(220, 0, 0) 
 End With 
End Sub
```


## See also


#### Concepts


[Shapes Collection Object](shapes-object-word.md)

