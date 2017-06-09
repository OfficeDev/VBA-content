---
title: DataLabel Object (PowerPoint)
keywords: vbapp10.chm696000
f1_keywords:
- vbapp10.chm696000
ms.prod: powerpoint
api_name:
- PowerPoint.DataLabel
ms.assetid: a17d23c5-0361-9129-28e5-b892f6966bda
ms.date: 06/08/2017
---


# DataLabel Object (PowerPoint)

Represents the data label on a chart point or trendline.


## Remarks

 On a series, the **DataLabel** object is a member of the **[DataLabels](datalabels-object-powerpoint.md)** collection. The **DataLabels** collection contains a **DataLabel** object for each point. For a series without definable points (such as an area series), the **DataLabels** collection contains a single **DataLabel** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use  **[DataLabels](series-datalabels-method-powerpoint.md)** ( _Index_ ), where _Index_ is the data label index number, to return a single **DataLabel** object. The following example sets the number format for the fifth data label in the first series of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).DataLabels(5).NumberFormat = "0.000"

    End If

End With


```




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use the  **[Point.DataLabel](point-datalabel-property-powerpoint.md)** property to return the **DataLabel** object for a single point. The following example turns on the data label for the second point in the first series of the first chart in the active document and sets the data label text to "Saturday."




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.SeriesCollection(1).Points(2)

            .HasDataLabel = True

            .DataLabel.Text = "Saturday"

        End With

    End If

End With


```




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

On a trendline, the  **[Trendline.DataLabel](trendline-datalabel-property-powerpoint.md)** property returns the text shown with the trendline. This can be the equation, the R-squared value, or both (if both are showing). The following example sets the trendline text for the first trendline in the first series of the first chart in the active document to show only the equation.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.SeriesCollection(1).Trendlines(1)

            .DisplayRSquared = False

            .DisplayEquation = True

        End With

    End If

End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

