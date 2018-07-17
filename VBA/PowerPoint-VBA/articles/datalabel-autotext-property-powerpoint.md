---
title: DataLabel.AutoText Property (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.DataLabel.AutoText
ms.assetid: f7e154ad-4f5f-0a3d-3fe5-c83994705cfb
ms.date: 06/08/2017
---


# DataLabel.AutoText Property (PowerPoint)

 **True** if the object automatically generates appropriate text based on context. Read/write **Boolean**.


## Syntax

 _expression_. **AutoText**

 _expression_ A variable that represents a **[DataLabel](datalabel-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the data labels for series one of the first chart in the active document to automatically generate appropriate text.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.SeriesCollection(1). _
            DataLabels.AutoText = True
    End If
End With
```


## See also


#### Concepts


[DataLabel Object](datalabel-object-powerpoint.md)

