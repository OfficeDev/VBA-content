---
title: ShapeNode.EditingType Property (Excel)
keywords: vbaxl10.chm111003
f1_keywords:
- vbaxl10.chm111003
ms.prod: excel
api_name:
- Excel.ShapeNode.EditingType
ms.assetid: 78a17ed7-7e30-d5f3-4af8-636d65079218
ms.date: 06/08/2017
---


# ShapeNode.EditingType Property (Excel)

If the specified node is a vertex, this property returns a value that indicates how changes made to the node affect the two segments connected to the node. Read-only  **[MsoEditingType](http://msdn.microsoft.com/library/5fe5c4f6-6467-c6a7-197c-ff700c384b92%28Office.15%29.aspx)** .


## Syntax

 _expression_ . **EditingType**

 _expression_ A variable that represents a **ShapeNode** object.


## Remarks

This property is read-only. Use the  **[SetEditingType](shapenodes-seteditingtype-method-excel.md)** method to set the value of this property.


## Example

This example changes all corner nodes to smooth nodes in shape three on  `myDocument`. Shape three must be a freeform drawing.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3).Nodes 
    For n = 1 to .Count 
        If .Item(n).EditingType = msoEditingCorner Then 
            .SetEditingType n, msoEditingSmooth 
        End If 
    Next 
End With
```


## See also


#### Concepts


[ShapeNode Object](shapenode-object-excel.md)

