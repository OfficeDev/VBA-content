---
title: Trendlines.Item Method (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Trendlines.Item
ms.assetid: ddda769f-ffc2-c03f-4087-755a5530f156
ms.date: 06/08/2017
---


# Trendlines.Item Method (PowerPoint)

Returns a single object from a collection.


## Syntax

 _expression_. **Item**( **_Index_** )

 _expression_ A variable that represents a **[Trendlines](trendlines-object-powerpoint.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional|**Variant**|The index number for the object.|

### Return Value

A  **[Trendline](trendline-object-powerpoint.md)** object that the collection contains.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the number of units that the trendline on the first chart in the active document extends forward and backward. The example should be run on a 2-D column chart that contains a single series with a trendline.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.SeriesCollection(1).Trendlines.Item(1)

            .Forward = 5

            .Backward = .5

        End With

    End If

End With
```


## See also


#### Concepts


[Trendlines Object](trendlines-object-powerpoint.md)

