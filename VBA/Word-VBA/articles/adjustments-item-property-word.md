---
title: Adjustments.Item Property (Word)
keywords: vbawd10.chm163840000
f1_keywords:
- vbawd10.chm163840000
ms.prod: word
api_name:
- Word.Adjustments.Item
ms.assetid: 10628688-e927-df50-a16a-e25878676c82
ms.date: 06/08/2017
---


# Adjustments.Item Property (Word)

Returns or sets the adjustment value specified by the  _Index_ argument. Read/write **Single** .


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ An expression that returns an **[Adjustments](adjustments-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The index number of the adjustment.|

## Remarks

For linear adjustments, an adjustment value of 0.0 generally corresponds to the left or top edge of the shape, and a value of 1.0 generally corresponds to the right or bottom edge of the shape. However, adjustments can pass beyond shape boundaries for some shapes. For radial adjustments, an adjustment value of 1.0 corresponds to the width of the shape. For angular adjustments, the adjustment value is specified in degrees. The  **Item** property applies only to shapes that have adjustments. AutoShapes and WordArt have up to eight adjustments.


## Example

This example adds two crosses to the active document and then sets the value for adjustment one (the only one for this type of AutoShape) on each cross.


```vb
Dim docActive As Document 
 
Set docActive = ActiveDocument 
With docActive.Shapes 
 .AddShape(msoShapeCross, _ 
 10, 10, 100, 100).Adjustments.Item(1) = 0.4 
 .AddShape(msoShapeCross, _ 
 150, 10, 100, 100).Adjustments.Item(1) = 0.2 
End With
```

This example has the same result as the previous example even though it doesn't explicitly use the  **Item** property.




```vb
Dim docActive As Document 
 
Set docActive = ActiveDocument 
With docActive.Shapes 
 .AddShape(msoShapeCross, _ 
 10, 10, 100, 100).Adjustments(1) = 0.4 
 .AddShape(msoShapeCross, _ 
 150, 10, 100, 100).Adjustments(1) = 0.2 
End With
```


## See also


#### Concepts


[Adjustments Object](adjustments-object-word.md)

