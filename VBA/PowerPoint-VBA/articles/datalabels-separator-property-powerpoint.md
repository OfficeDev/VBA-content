---
title: DataLabels.Separator Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.DataLabels.Separator
ms.assetid: e0bc6147-61c8-8df9-ff42-591f60c5b7f5
ms.date: 06/08/2017
---


# DataLabels.Separator Property (PowerPoint)

Sets or returns the separator for the data labels on a chart. Read/write  **Variant**.


## Syntax

 _expression_. **Separator**

 _expression_ A variable that represents a **[DataLabels](datalabels-object-powerpoint.md)** object.


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


[DataLabels Object](datalabels-object-powerpoint.md)

