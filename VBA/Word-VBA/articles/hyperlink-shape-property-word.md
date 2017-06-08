---
title: Hyperlink.Shape Property (Word)
keywords: vbawd10.chm161285103
f1_keywords:
- vbawd10.chm161285103
ms.prod: word
api_name:
- Word.Hyperlink.Shape
ms.assetid: bee91eb6-fc38-e2b9-ca90-e9a34062c9f5
ms.date: 06/08/2017
---


# Hyperlink.Shape Property (Word)

Returns a  **Shape** object for the specified hyperlink or diagram node.


## Syntax

 _expression_ . **Shape**

 _expression_ Required. A variable that represents a **[Hyperlink](hyperlink-object-word.md)** object.


## Remarks

If a hyperlink isn't represented by a shape, an error occurs.


## Example

This example changes the fill color for the shape that represents the first hyperlink in the active document. For this example to work, the hyperlink must be represented by a shape.


```vb
ActiveDocument.Hyperlinks(1).Shape.Fill.ForeColor.RGB = _ 
 RGB(255, 255, 0)
```


## See also


#### Concepts


[Hyperlink Object](hyperlink-object-word.md)

