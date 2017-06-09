---
title: GroupShapes.Range Property (Excel)
keywords: vbaxl10.chm642077
f1_keywords:
- vbaxl10.chm642077
ms.prod: excel
api_name:
- Excel.GroupShapes.Range
ms.assetid: b02b1915-2cd2-353b-0243-a5d60470e897
ms.date: 06/08/2017
---


# GroupShapes.Range Property (Excel)

Returns a  **[ShapeRange](shaperange-object-excel.md)** object that represents a subset of the shapes in a **[Shapes](shapes-object-excel.md)** collection.


## Syntax

 _expression_ . **Range**( **_Index_** )

 _expression_ A variable that represents a **GroupShapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The individual shapes to be included in the range. Can be an integer that specifies the index number of the shape, a string that specifies the name of the shape, or an array that contains either integers or strings.|

## Remarks

Although you can use the  **[Range](groupshapes-range-property-excel.md)** property to return any number of shapes, it is simpler to use the **[Item](groupshapes-item-method-excel.md)** method if you want to return only a single member of the collection. For example, `Shapes(1)` is simpler than `Shapes.Range(1)`.


## Example

This example sets the fill pattern for shapes one and three on  `myDocument`.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes.Range(Array(1, 3)) _ 
 .Fill.Patterned msoPatternHorizontalBrick
```

To specify an array of integers or strings for  **Index** , you can use the **Array** function. For example, the following instruction returns two shapes specified by name.




```vb
Dim arShapes() As Variant 
Dim objRange As Object 
arShapes = Array("Oval 4", "Rectangle 5") 
Set objRange = ActiveSheet.Shapes.Range(arShapes) 
 
```

In Microsoft Excel, you cannot use this property to return a  **ShapeRange** object containing all the **Shape** objects on a worksheet. Instead, use the following code:




```vb
Worksheets(1).Shapes.SelectAll ' select all shapes 
set sr = Selection.ShapeRange ' create ShapeRange 
 
```

This example sets the fill pattern for the shapes named "Oval 4" and "Rectangle 5" on  `myDocument`.




```vb
Dim arShapes() As Variant 
Dim objRange As Object 
Set myDocument = Worksheets(1) 
arShapes = Array("Oval 4", "Rectangle 5") 
Set objRange = myDocument.Shapes.Range(arShapes) 
objRange.Fill.Patterned msoPatternHorizontalBrick
```

This example sets the fill pattern for shape one on  `myDocument`.




```vb
Set myDocument = Worksheets(1) 
Set myRange = myDocument.Shapes.Range(1) 
myRange.Fill.Patterned msoPatternHorizontalBrick
```

This example creates an array that contains all the AutoShapes on  `myDocument`, uses that array to define a shape range, and then distributes all the shapes in that range horizontally.




```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes 
 numShapes = .Count 
 If numShapes > 1 Then 
 numAutoShapes = 1 
 ReDim autoShpArray(1 To numShapes) 
 For i = 1 To numShapes 
 If .Item(i).Type = msoAutoShape Then 
 autoShpArray(numAutoShapes) = .Item(i).Name 
 numAutoShapes = numAutoShapes + 1 
 End If 
 Next 
 If numAutoShapes > 1 Then 
 ReDim Preserve autoShpArray(1 To numAutoShapes) 
 Set asRange = .Range(autoShpArray) 
 asRange.Distribute msoDistributeHorizontally, False 
 End If 
 End If 
End With
```


## See also


#### Concepts


[GroupShapes Object](groupshapes-object-excel.md)

