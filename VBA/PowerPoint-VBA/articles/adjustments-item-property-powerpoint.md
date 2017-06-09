---
title: Adjustments.Item Property (PowerPoint)
keywords: vbapp10.chm550003
f1_keywords:
- vbapp10.chm550003
ms.prod: powerpoint
api_name:
- PowerPoint.Adjustments.Item
ms.assetid: 54cc6850-0fe8-8887-2acc-dc91085b7451
ms.date: 06/08/2017
---


# Adjustments.Item Property (PowerPoint)

Returns or sets the adjustment value specified by the Index argument. Read/write.


## Syntax

 _expression_. **Item**( **_Index_** )

 _expression_ A variable that represents an **Adjustments** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Long**|The index number of the adjustment.|

### Return Value

Single


## Remarks

For linear adjustments, an adjustment value of 0.0 generally corresponds to the left or top edge of the shape, and a value of 1.0 generally corresponds to the right or bottom edge of the shape. However, adjustments can pass beyond shape boundaries for some shapes. For radial adjustments, an adjustment value of 1.0 corresponds to the width of the shape. For angular adjustments, the adjustment value is specified in degrees. The  **Item** property applies only to shapes that have adjustments.

AutoShapes, connectors, and WordArt objects can have up to eight adjustments.


## Example

This example adds two crosses to  `myDocument` and then sets the value for adjustment one (the only one on this type of AutoShape) on each cross.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes
    .AddShape(msoShapeCross, 10, 10, 100, 100) _
        .Adjustments.Item(1) = 0.4

    .AddShape(msoShapeCross, 150, 10, 100, 100) _
        .Adjustments.Item(1) = 0.2
End With
```

This example has the same result as the previous example even though it doesn't explicitly use the  **Item** property.




```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes
    .AddShape(msoShapeCross, 10, 10, 100, 100) _
        .Adjustments(1) = 0.4

    .AddShape(msoShapeCross, 150, 10, 100, 100) _
        .Adjustments(1) = 0.2
End With
```


## See also


#### Concepts


[Adjustments Object](adjustments-object-powerpoint.md)

