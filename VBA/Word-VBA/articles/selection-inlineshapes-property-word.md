---
title: Selection.InlineShapes Property (Word)
keywords: vbawd10.chm158663067
f1_keywords:
- vbawd10.chm158663067
ms.prod: word
api_name:
- Word.Selection.InlineShapes
ms.assetid: 2fbbf39c-b70e-e332-2547-089166e718ca
ms.date: 06/08/2017
---


# Selection.InlineShapes Property (Word)

Returns an  **[InlineShapes](inlineshapes-object-word.md)** collection that represents all the **InlineShape** objects in a selection. Read-only.


## Syntax

 _expression_ . **InlineShapes**

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example displays the number of shapes and inline shapes in the active document.


```vb
Set doc = ActiveDocument 
Msgbox "InlineShape = " &; doc.InlineShapes.Count &; _ 
 vbCr &; "Shapes = " &; doc.Shapes.Count
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

