---
title: Axis.MajorUnit Property (PowerPoint)
keywords: vbapp10.chm682013
f1_keywords:
- vbapp10.chm682013
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.MajorUnit
ms.assetid: 5f88f369-e999-b947-c47f-5413e349d192
ms.date: 06/08/2017
---


# Axis.MajorUnit Property (PowerPoint)

Returns or sets the major units for the value axis. Read/write  **Double**.


## Syntax

 _expression_. **MajorUnit**

 _expression_ A variable that represents an **[Axis](axis-object-powerpoint.md)** object.


## Remarks

Setting this property sets the  **[MajorUnitIsAuto](axis-majorunitisauto-property-powerpoint.md)** property to **False**.

Use the  **[TickMarkSpacing](axis-tickmarkspacing-property-powerpoint.md)** property to set tick mark spacing on the category axis.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the major and minor units for the value axis of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Axes(xlValue)

            .MajorUnit = 100

            .MinorUnit = 20

        End With

    End If

End With
```


## See also


#### Concepts


[Axis Object](axis-object-powerpoint.md)

