---
title: Selection.ShapeRange Property (Word)
keywords: vbawd10.chm158663660
f1_keywords:
- vbawd10.chm158663660
ms.prod: word
api_name:
- Word.Selection.ShapeRange
ms.assetid: b327da9a-8858-1ec1-8a0d-cb36bd44fede
ms.date: 06/08/2017
---


# Selection.ShapeRange Property (Word)

Returns a  **[ShapeRange](shaperange-object-word.md)** collection that represents all the **Shape** objects in the selection. Read-only.


## Syntax

 _expression_ . **ShapeRange**

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

The shape range can contain drawings, shapes, pictures, OLE objects, ActiveX controls, text objects, and callouts. 


## Example

The following example applies shadow formatting to all the shapes in the selection.


```
Selection.ShapeRange.Shadow.Type = msoShadow6
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

