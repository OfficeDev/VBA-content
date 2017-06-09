---
title: CatalogMergeShapes.Range Method (Publisher)
keywords: vbapb10.chm8388612
f1_keywords:
- vbapb10.chm8388612
ms.prod: publisher
api_name:
- Publisher.CatalogMergeShapes.Range
ms.assetid: e92dcac4-4694-8a22-61da-09fcd98c72ce
ms.date: 06/08/2017
---


# CatalogMergeShapes.Range Method (Publisher)

Returns a  **[ShapeRange](shaperange-object-publisher.md)** object that represents a subset of the shapes in a **Shapes** collection.


## Syntax

 _expression_. **Range**( **_Index_**)

 _expression_A variable that represents a  **CatalogMergeShapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Variant**|The individual shapes that are to be included in the range. Can be an integer that specifies the index number of the shape, a string that specifies the name of the shape, or an array that contains either integers or strings. If Index is omitted, the  **Range** method returns all the objects in the specified collection.|

### Return Value

ShapeRange


## Example

To specify an array of integers or strings for  **_Index_**, you can use the  **Array** function. For example, the following instruction returns two shapes specified by name.


```vb
Dim arrShapes As Variant 
Dim shpRange As ShapeRange 
 
Set arrShapes = Array("Oval 4", "Rectangle 5") 
Set shpRange = ActiveDocument.Pages(1) _ 
 .Shapes.Range(arrShapes)
```

This example sets the fill pattern for shapes one and three on the active publication.




```vb
ActiveDocument.Pages(1).Shapes.Range(Array(1, 3)).Fill _ 
 .Patterned msoPatternHorizontalBrick
```




```

```

This example sets the fill pattern for the shapes named "Oval 4" and "Rectangle 5" on the first page.




```vb
Dim arrShapes As Variant 
Dim shpRange As ShapeRange 
 
arrShapes = Array("Oval 4", "Rectangle 5") 
 
Set shpRange = ActiveDocument.Pages(1).Shapes.Range(arrShapes) 
 
shpRange.Fill.Patterned msoPatternHorizontalBrick
```

This example sets the fill pattern for all shapes on the first page.




```vb
ActiveDocument.Pages(1).Shapes _ 
 .Range.Fill.Patterned msoPatternHorizontalBrick
```

This example sets the fill pattern for shape one on the first page.




```vb
Dim shpRange As ShapeRange 
 
Set shpRange = ActiveDocument.Pages(1).Shapes.Range(1) 
 
shpRange.Fill.Patterned msoPatternHorizontalBrick
```

This example creates an array that contains all the AutoShapes on the first page, uses that array to define a shape range, and then distributes all the shapes in that range horizontally.




```vb
Dim numShapes As Long 
Dim numAutoShapes As Long 
Dim autoShpArray As Variant 
Dim intLoop As Integer 
Dim shpRange As ShapeRange 
 
With ActiveDocument.Pages(1).Shapes 
 
 numShapes = .Count 
 If numShapes > 1 Then 
 
 numAutoShapes = 0 
 ReDim autoShpArray(1 To numShapes) 
 
 For intLoop = 1 To numShapes 
 If .Item(intLoop).Type = msoAutoShape Then 
 numAutoShapes = numAutoShapes + 1 
 autoShpArray(numAutoShapes) = .Item(intLoop).Name 
 End If 
 Next 
 
 If numAutoShapes > 1 Then 
 ReDim Preserve autoShpArray(1 To numAutoShapes) 
 Set shpRange = .Range(autoShpArray) 
 shpRange.Distribute _ 
 DistributeCmd:=msoDistributeHorizontally, _ 
 RelativeTo:=False 
 End If 
 
 End If 
 
End With
```


