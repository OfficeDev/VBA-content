---
title: Document.InlineShapes Property (Word)
keywords: vbawd10.chm158007364
f1_keywords:
- vbawd10.chm158007364
ms.prod: word
api_name:
- Word.Document.InlineShapes
ms.assetid: 049510b5-cdb3-74e8-783a-4c8fa809b876
ms.date: 06/08/2017
---


# Document.InlineShapes Property (Word)

Returns an  **[InlineShapes](document-inlineshapes-property-word.md)** collection that represents all the **[InlineShape](inlineshape-object-word.md)** objects in a document. Read-only.


## Syntax

 _expression_ . **InlineShapes**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


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


[Document Object](document-object-word.md)

