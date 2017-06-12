---
title: Shape.Ungroup Method (Word)
keywords: vbawd10.chm161480727
f1_keywords:
- vbawd10.chm161480727
ms.prod: word
api_name:
- Word.Shape.Ungroup
ms.assetid: 0e8ead12-19a7-4caf-696e-38509e30148d
ms.date: 06/08/2017
---


# Shape.Ungroup Method (Word)

Ungroups any grouped shapes in the specified shape.


## Syntax

 _expression_ . **Ungroup**

 _expression_ Required. A variable that represents a **[Shape](shape-object-word.md)** object.


### Return Value

ShapeRange


## Remarks

This method sisassembles pictures and OLE objects within the specified shapeand returns the ungrouped shapes as a single  **[ShapeRange](shaperange-object-word.md)** object.

Because a group of shapes is treated as a single object, grouping and ungrouping shapes changes the number of items in the  **Shapes** collection and changes the index numbers of items that come after the affected items in the collection.


## Example

This example ungroups any grouped shapes and disassembles any pictures or OLE objects on  _myDocument_ .


```vb
Set myDocument = ActiveDocument 
For Each s In myDocument.Shapes 
 s.Ungroup 
Next
```

This example ungroups any grouped shapes on  _myDocument_ without disassembling pictures or OLE objects on the document.




```vb
Set myDocument = ActiveDocument 
For Each s In myDocument.Shapes 
 If s.Type = msoGroup Then s.Ungroup 
Next
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

