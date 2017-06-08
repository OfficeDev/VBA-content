---
title: DataLabels.Item Method (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.DataLabels.Item
ms.assetid: 233cb110-f20c-4e68-9033-f9c2073ac061
ms.date: 06/08/2017
---


# DataLabels.Item Method (PowerPoint)

Returns a single object from a collection.


## Syntax

 _expression_. **Item**( **_Index_** )

 _expression_ A variable that represents a **[DataLabels](datalabels-object-powerpoint.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**|The index number for the object.|

### Return Value

A  **[DataLabel](datalabel-object-powerpoint.md)** object contained by the collection.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the number format for the fifth data label in the first series for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.SeriesCollection(1).DataLabels.Item(5). _
            NumberFormat = "0.000"
    End If
End With


```


## See also


#### Concepts


[DataLabels Object](datalabels-object-powerpoint.md)

