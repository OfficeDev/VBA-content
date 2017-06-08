---
title: ShapeNode Object (Excel)
keywords: vbaxl10.chm111000
f1_keywords:
- vbaxl10.chm111000
ms.prod: excel
api_name:
- Excel.ShapeNode
ms.assetid: c8b60d74-f11f-1659-30a3-6e180eb8bd58
ms.date: 06/08/2017
---


# ShapeNode Object (Excel)

Represents the geometry and the geometry-editing properties of the nodes in a user-defined freeform.


## Remarks

 Nodes include the vertices between the segments of the freeform and the control points for curved segments. The **ShapeNode** object is a member of the **[ShapeNodes](shapenodes-object-excel.md)** collection. The **ShapeNodes** collection contains all the nodes in a freeform.


## Example

Use  **[Nodes](shape-nodes-property-excel.md)** ( _index_ ), where _index_ is the node index number, to return a single **ShapeNode** object. If node one in shape three on _myDocument_ is a corner point, the following example makes it a smooth point. For this example to work, shape three must be a freeform.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3) 
 If .Nodes(1).EditingType = msoEditingCorner Then 
 .Nodes.SetEditingType 1, msoEditingSmooth 
 End If 
End With
```


## See also


#### Other resources



[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

