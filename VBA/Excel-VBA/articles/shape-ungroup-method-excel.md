---
title: Shape.Ungroup Method (Excel)
keywords: vbaxl10.chm636087
f1_keywords:
- vbaxl10.chm636087
ms.prod: excel
api_name:
- Excel.Shape.Ungroup
ms.assetid: 678ff982-25c7-cbaa-7cc5-011b53ecf6b6
ms.date: 06/08/2017
---


# Shape.Ungroup Method (Excel)

Ungroups any grouped shapes in the specified shape or range of shapes. Disassembles pictures and OLE objects within the specified shape or range of shapes.


## Syntax

 _expression_ . **Ungroup**

 _expression_ A variable that represents a **Shape** object.


### Return Value

A  **[ShapeRange](shaperange-object-excel.md)** object that represents the ungrouped shapes.


## Remarks

Because a group of shapes is treated as a single object, grouping and ungrouping shapes changes the number of items in the  **Shapes** collection and changes the index numbers of items that come after the affected items in the collection.


## Example

This example ungroups any grouped shapes and disassembles any pictures or OLE objects on  `myDocument`.


```vb
Set myDocument = Worksheets(1) 
For Each s In myDocument.Shapes 
 s.Ungroup 
Next
```

This example ungroups any grouped shapes on  `myDocument` without disassembling pictures or OLE objects on the document.




```vb
Set myDocument = Worksheets(1) 
For Each s In myDocument.Shapes 
 If s.Type = msoGroup Then s.Ungroup
```


## See also


#### Concepts


[Shape Object](shape-object-excel.md)

