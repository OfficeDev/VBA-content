---
title: Axis.HasTitle Property (PowerPoint)
keywords: vbapp10.chm682010
f1_keywords:
- vbapp10.chm682010
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.HasTitle
ms.assetid: 04f9e10a-f323-a905-e09c-e9bb3222a80c
ms.date: 06/08/2017
---


# Axis.HasTitle Property (PowerPoint)

 **True** if the axis or chart has a visible title. Read/write **Boolean**.


## Syntax

 _expression_. **HasTitle**

 _expression_ A variable that represents an **[Axis](axis-object-powerpoint.md)** object.


## Remarks

An axis title is represented by an  **[AxisTitle](axistitle-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example adds an axis label to the category axis for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Axis(xlCategory)

            .HasTitle = True

            .AxisTitle.Text = "July Sales"

        End With

    End If

End With
```


## See also


#### Concepts


[Axis Object](axis-object-powerpoint.md)

