---
title: DataLabels.ShowCategoryName Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.DataLabels.ShowCategoryName
ms.assetid: 0869b709-e09d-2c55-4d74-c4a0d130a551
ms.date: 06/08/2017
---


# DataLabels.ShowCategoryName Property (PowerPoint)

 **True** to display the category name for the data labels on a chart. **False** to hide the name. Read/write **Boolean**.


## Syntax

 _expression_. **ShowCategoryName**

 _expression_ A variable that represents a **[DataLabels](datalabels-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example shows the category name for the data labels of the first series on the first chart.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then
        .Chart.SeriesCollection(1).DataLabels. _
            ShowCategoryName = True
    End If

End With
```


## See also


#### Concepts


[DataLabels Object](datalabels-object-powerpoint.md)

