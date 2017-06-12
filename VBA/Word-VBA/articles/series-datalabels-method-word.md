---
title: Series.DataLabels Method (Word)
keywords: vbawd10.chm123732125
f1_keywords:
- vbawd10.chm123732125
ms.prod: word
api_name:
- Word.Series.DataLabels
ms.assetid: 6dc09ab8-efa4-a5cb-2372-989bf6e3769c
ms.date: 06/08/2017
---


# Series.DataLabels Method (Word)

Returns an object that represents either a single data label (a  **[DataLabel](datalabel-object-word.md)** object) or a collection of all the data labels for the series (a **[DataLabels](datalabels-object-word.md)** collection).


## Syntax

 _expression_ . **DataLabels**( **_Index_** )

 _expression_ A variable that represents a **[Series](series-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The number of the data label.|

### Return Value

An object that represents either a single data label (a  **DataLabel** object) or a collection of all the data labels for the series (a **DataLabels** collection).


## Remarks

If the series has the  **Show Value** option turned on for the data labels, the returned collection can contain up to one label for each point. Data labels can be turned on or off for individual points in the series.

If the series is on an area chart and has the  **Show Label** option turned on for the data labels, the returned collection contains only a single label, which is the label for the area series.


## Example

The following example sets the data labels for series one for the first chart in the active document to show their key, assuming that their values are visible when the example runs.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.SeriesCollection(1) 
 .HasDataLabels = True 
 With .DataLabels 
 .ShowLegendKey = True 
 .Type = xlValue 
 End With 
 End With 
 End If 
End With
```


## See also


#### Concepts


[Series Object](series-object-word.md)

