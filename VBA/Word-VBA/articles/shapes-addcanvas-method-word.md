---
title: Shapes.AddCanvas Method (Word)
keywords: vbawd10.chm161415193
f1_keywords:
- vbawd10.chm161415193
ms.prod: word
api_name:
- Word.Shapes.AddCanvas
ms.assetid: ff6da70f-f6ce-83f8-8e30-95b50a1f4e4f
ms.date: 06/08/2017
---


# Shapes.AddCanvas Method (Word)

Adds a drawing canvas to a document. Returns a  **[Shape](shape-object-word.md)** object that represents the drawing canvas and adds it to the **Shapes** collection.


## Syntax

 _expression_ . **AddCanvas**( **_Left_** , **_Top_** , **_Width_** , **_Height_** , **_Anchor_** )

 _expression_ Required. A variable that represents a **[Shapes](shapes-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Left_|Required| **Single**|The position, in points, of the left edge of the drawing canvas, relative to the anchor.|
| _Top_|Required| **Single**|The position, in points, of the top edge of the drawing canvas, relative to the anchor.|
| _Width_|Required| **Single**|The width, in points, of the drawing canvas.|
| _Height_|Required| **Single**|The height, in points, of the drawing canvas.|
| _Anchor_|Optional| **Variant**|A  **[Range](range-object-word.md)** object that represents the text to which the canvas is bound. If Anchor is specified, the anchor is positioned at the beginning of the first paragraph in the anchoring range. If this argument is omitted, the anchoring range is selected automatically and the canvas is positioned relative to the top and left edges of the page.|

### Return Value

Shape


## Example

The following example adds a drawing canvas to a new document and formats the drawing canvas so it is inline with the text; then adds two shapes to the canvas, and formats the fill and line properties.


```vb
Sub AddInlineCanvas() 
 Dim docNew As Document 
 Dim shpCanvas As Shape 
 
 Set docNew = Documents.Add 
 
 'Add a drawing canvas to the new document 
 Set shpCanvas = docNew.Shapes.AddCanvas( _ 
 Left:=150, Top:=150, Width:=70, Height:=70) 
 shpCanvas.WrapFormat.Type = wdWrapInline 
 
 'Add shapes to drawing canvas 
 With shpCanvas.CanvasItems 
 .AddShape msoShapeHeart, Left:=10, _ 
 Top:=10, Width:=50, Height:=60 
 .AddLine BeginX:=0, BeginY:=0, _ 
 EndX:=70, EndY:=70 
 End With 
 With shpCanvas 
 .CanvasItems(1).Fill.ForeColor _ 
 .RGB = RGB(Red:=255, Green:=0, Blue:=0) 
 .CanvasItems(2).Line _ 
 .EndArrowheadStyle = msoArrowheadTriangle 
 End With 
End Sub
```


## See also


#### Concepts


[Shapes Collection Object](shapes-object-word.md)

