---
title: Series.Points Method (PowerPoint)
keywords: vbapp10.chm65606
f1_keywords:
- vbapp10.chm65606
ms.prod: powerpoint
api_name:
- PowerPoint.Series.Points
ms.assetid: 53bec845-d3a0-fdce-921b-66d2d4e1eb59
ms.date: 06/08/2017
---


# Series.Points Method (PowerPoint)

Returns a collection of all the points in the series.


## Syntax

 _expression_. **Points**( **_Index_** )

 _expression_ A variable that represents a **[Series](series-object-powerpoint.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional|**Variant**|The name or number of the point.|

### Return Value

A  **[Points](points-object-powerpoint.md)** object that represents all the points in the series.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example applies a data label to point one in series one of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).Points(1).ApplyDataLabels

    End If

End With
```


## See also


#### Concepts


[Series Object](series-object-powerpoint.md)

