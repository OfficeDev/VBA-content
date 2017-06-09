---
title: SeriesCollection.Item Method (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.SeriesCollection.Item
ms.assetid: ae34ad0d-1b0a-decb-24e8-3d1c51652f72
ms.date: 06/08/2017
---


# SeriesCollection.Item Method (PowerPoint)

Returns a single object from a collection.


## Syntax

 _expression_. **Item**( **_Index_** )

 _expression_ A variable that represents a **[SeriesCollection](seriescollection-object-powerpoint.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**|The name or index number for the object.|

### Return Value

A  **[Series](series-object-powerpoint.md)** object contained by the collection.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the number of units that the trendline on the first chart in the active document extends forward and backward. The example should be run on a 2-D column chart that contains a single series with a trendline.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.SeriesCollection.Item(1).Trendlines.Item(1)

            .Forward = 5

            .Backward = .5

        End With

    End If

End With


```


## See also


#### Concepts


[SeriesCollection Object](seriescollection-object-powerpoint.md)

