---
title: Points.Item Method (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Points.Item
ms.assetid: d3a6b3cf-3fbb-1e0f-b9cf-0b707839de67
ms.date: 06/08/2017
---


# Points.Item Method (PowerPoint)

Returns a single object from a collection.


## Syntax

 _expression_. **Item**( **_Index_** )

 _expression_ A variable that represents a **[Points](points-object-powerpoint.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Long**|The index number for the object.|

### Return Value

A  **[Point](point-object-powerpoint.md)** object that the collection contains.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the marker style for the third point in series one in embedded chart one on worksheet one. The specified series must be a 2-D line, scatter, or radar series.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.SeriesCollection(1).Points.Item(3). _
            MarkerStyle = xlDiamond
    End If
End With
```


## See also


#### Concepts


[Points Object](points-object-powerpoint.md)

