---
title: ShapeRange.Distribute Method (Word)
keywords: vbawd10.chm162856973
f1_keywords:
- vbawd10.chm162856973
ms.prod: word
api_name:
- Word.ShapeRange.Distribute
ms.assetid: fae7b87a-9542-7018-15fb-a4e4efee4c9b
ms.date: 06/08/2017
---


# ShapeRange.Distribute Method (Word)

Evenly distributes the shapes in the specified range of shapes. .


## Syntax

 _expression_ . **Distribute**( **_Distribute_** , **_RelativeTo_** )

 _expression_ Required. A variable that represents a **[ShapeRange](shaperange-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Distribute_|Required| **MsoDistributeCmd**|Specifies whether to distribute shapes horizontally or vertically.|
| _RelativeTo_|Required| **Long**| **True** to distribute the shapes evenly over the entire horizontal or vertical space on the page. **False** to distribute them within the horizontal or vertical space that the range of shapes originally occupies.|

## Remarks

You can specify whether you want to distribute the shapes horizontally or vertically and whether you want to distribute them over the entire page or just over the space they originally occupy.


## Example

This example defines a shape range that contains all the AutoShapes on the active document and then horizontally distributes the shapes in this range.


```vb
With ActiveDocument.Shapes 
 numShapes = .Count 
 If numShapes > 1 Then 
 numAutoShapes = 0 
 ReDim autoShpArray(1 To numShapes) 
 For i = 1 To numShapes 
 If .Item(i).Type = msoAutoShape Then 
 numAutoShapes = numAutoShapes + 1 
 autoShpArray(numAutoShapes) = .Item(i).Name 
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


[ShapeRange Collection Object](shaperange-object-word.md)

