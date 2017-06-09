---
title: ShapeNode.EditingType Property (Word)
keywords: vbawd10.chm164429924
f1_keywords:
- vbawd10.chm164429924
ms.prod: word
api_name:
- Word.ShapeNode.EditingType
ms.assetid: ac490e3c-3938-a1db-50b5-ec667061f711
ms.date: 06/08/2017
---


# ShapeNode.EditingType Property (Word)

If the specified node is a vertex, this property returns a value that indicates how changes made to the node affect the two segments connected to the node. Read-only  **MsoEditingType** . .


## Syntax

 _expression_ . **EditingType**

 _expression_ Required. A variable that represents a **[ShapeNode](shapenode-object-word.md)** object.


## Remarks

If the node is a control point for a curved segment, this property returns the editing type of the adjacent vertex. This property is read-only. Use the  **SetEditingType** method to set the value of this property.


## Example

This example changes all corner nodes to smooth nodes in the third shape on the active document. The third shape must be a freeform drawing.


```vb
Dim docActive As Document 
Dim intCount As Integer 
 
Set docActive = ActiveDocument 
 
With docActive.Shapes(3).Nodes 
 For intCount = 1 to .Count 
 If .Item(intCount).EditingType = msoEditingCorner Then 
 .SetEditingType intCount, msoEditingSmooth 
 End If 
 Next 
End With
```


## See also


#### Concepts


[ShapeNode Object](shapenode-object-word.md)

