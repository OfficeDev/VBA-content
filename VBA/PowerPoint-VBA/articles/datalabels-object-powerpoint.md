---
title: DataLabels Object (PowerPoint)
keywords: vbapp10.chm697000
f1_keywords:
- vbapp10.chm697000
ms.prod: powerpoint
api_name:
- PowerPoint.DataLabels
ms.assetid: a0d0b0ec-6a12-9a5c-1026-1e1d85e488fa
ms.date: 06/08/2017
---


# DataLabels Object (PowerPoint)

A collection of all the  **[DataLabel](datalabel-object-powerpoint.md)** objects for the specified series.


## Remarks

 Each **DataLabel** object represents a data label for a point or trendline. For a series without definable points (such as an area series), the **DataLabels** collection contains a single data label.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use the  **[DataLabels](series-datalabels-method-powerpoint.md)** method to return the **DataLabels** collection. The following example sets the number format for data labels on the first series of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With Chart.SeriesCollection(1)

            .HasDataLabels = True

            .DataLabels.NumberFormat = "##.##"

        End With

    End If

End With


```




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use  **[DataLabels](series-datalabels-method-powerpoint.md)** ( _Index_ ), where _Index_ is the data label index number, to return a single **DataLabel** object. The following example sets the number format for the fifth data label in the first series of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With Chart.SeriesCollection(1).DataLabels(5)

            .NumberFormat = "0.000"

        End With

    End If

End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

