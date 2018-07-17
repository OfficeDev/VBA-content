---
title: ShapeNodes Object (Excel)
keywords: vbaxl10.chm112000
f1_keywords:
- vbaxl10.chm112000
ms.prod: excel
api_name:
- Excel.ShapeNodes
ms.assetid: 663721f1-8bd0-dd21-2362-fea2da3988bf
ms.date: 06/08/2017
---


# ShapeNodes Object (Excel)

A collection of all the  **[ShapeNode](shapenode-object-excel.md)** objects in the specified freeform.


## Remarks

 Each **ShapeNode** object represents either a node between segments in a freeform or a control point for a curved segment of a freeform. You can create a freeform manually or by using the **[BuildFreeform](shapes-buildfreeform-method-excel.md)** and **[ConvertToShape](freeformbuilder-converttoshape-method-excel.md)** methods.


## Example

Use the  **[Nodes](shape-nodes-property-excel.md)** property to return the **[ShapeNodes](shapenodes-object-excel.md)** collection. The following example deletes node four in shape three on _myDocument_ . For this example to work, shape three must be a freeform with at least four nodes.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes(3).Nodes.Delete 4
```

Use the  **[Insert](shapenodes-insert-method-excel.md)** method to create a new node and add it to the **ShapeNodes** collection. The following example adds a smooth node with a curved segment after node four in shape three on _myDocument_ . For this example to work, shape three must be a freeform with at least four nodes.




```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3).Nodes 
 .Insert 4, msoSegmentCurve, msoEditingSmooth, 210, 100 
End With
```

Use  **Nodes** ( _index_ ), where _index_ is the node index number, to return a single **ShapeNode** object. If node one in shape three on _myDocument_ is a corner point, the following example makes it a smooth point. For this example to work, shape three must be a freeform.




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


