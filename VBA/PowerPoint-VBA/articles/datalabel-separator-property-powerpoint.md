---
title: DataLabel.Separator Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.DataLabel.Separator
ms.assetid: 16613cac-f04d-13fe-56e5-bb6b6c9473b3
ms.date: 06/08/2017
---


# DataLabel.Separator Property (PowerPoint)

Returns or sets the separator used for the data labels on a chart. Read/write  **Variant**.


## Syntax

 _expression_. **Separator**

 _expression_ A variable that represents a **[DataLabel](datalabel-object-powerpoint.md)** object.


## Remarks

If you use a string, you will get a string as the separator. If you use  **xlDataLabelSeparatorDefault** (= 1), you will get the default data label separator, which is either a comma or a newline character, depending on the data label.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the data label separator for the first series on the first chart in the active document to a semicolon.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.SeriesCollection(1). _
            DataLabels.Separator = ";"
    End If
End With
```


## See also


#### Concepts


[DataLabel Object](datalabel-object-powerpoint.md)

