---
title: Series.HasDataLabels Property (PowerPoint)
keywords: vbapp10.chm65614
f1_keywords:
- vbapp10.chm65614
ms.prod: powerpoint
api_name:
- PowerPoint.Series.HasDataLabels
ms.assetid: b0b9bd37-7416-9903-d656-c4e468a9e481
ms.date: 06/08/2017
---


# Series.HasDataLabels Property (PowerPoint)

 **True** if the series has data labels. Read/write **Boolean**.


## Syntax

 _expression_. **HasDataLabels**

 _expression_ A variable that represents a **[Series](series-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example enables data labels for series three of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.SeriesCollection(3)

            .HasDataLabels = True

            .ApplyDataLabels Type:=xlValue

        End With

    End If

End With
```


## See also


#### Concepts


[Series Object](series-object-powerpoint.md)

