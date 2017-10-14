---
title: Axis.DisplayUnitLabel Property (PowerPoint)
keywords: vbapp10.chm682045
f1_keywords:
- vbapp10.chm682045
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.DisplayUnitLabel
ms.assetid: 75b01ce4-8edd-bbaa-d0fb-2d36c96b4da6
ms.date: 06/08/2017
---


# Axis.DisplayUnitLabel Property (PowerPoint)

Returns the  **[DisplayUnitLabel](displayunitlabel-object-powerpoint.md)** object for the specified axis. Returns **null** if the **[HasDisplayUnitLabel](axis-hasdisplayunitlabel-property-powerpoint.md)** property is set to **False**. Read-only.


## Syntax

 _expression_. **DisplayUnitLabel**

 _expression_ A variable that represents an **[Axis](axis-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the label caption to "Millions" for the value axis of the first chart in the active document, and then it turns off automatic font scaling.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Axes(xlValue).DisplayUnitLabel

            .Caption = "Millions"

            .AutoScaleFont = False

        End With

    End If

End With
```


## See also


#### Concepts


[Axis Object](axis-object-powerpoint.md)

