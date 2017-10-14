---
title: CalloutFormat.Drop Property (PowerPoint)
keywords: vbapp10.chm559011
f1_keywords:
- vbapp10.chm559011
ms.prod: powerpoint
api_name:
- PowerPoint.CalloutFormat.Drop
ms.assetid: 634bc753-2960-b699-535e-93c66fce280d
ms.date: 06/08/2017
---


# CalloutFormat.Drop Property (PowerPoint)

For callouts with an explicitly set drop value, this property returns the vertical distance (in points) from the edge of the text bounding box to the place where the callout line attaches to the text box. Read-only.


## Syntax

 _expression_. **Drop**

 _expression_ A variable that represents a **CalloutFormat** object.


### Return Value

Single


## Remarks

The distance is measured from the top of the text box unless the  **AutoAttach** property is set to **True** and the text box is to the left of the origin of the callout line (the place that the callout points to). In this case the drop distance is measured from the bottom of the text box.

Use the  **[CustomDrop](calloutformat-customdrop-method-powerpoint.md)** method to set the value of this property.

The value of this property accurately reflects the position of the callout line attachment to the text box only if the callout has an explicitly set drop value — that is, if the value of the  **[DropType](calloutformat-droptype-property-powerpoint.md)** property is **msoCalloutDropCustom**.


## Example

This example replaces the custom drop for shape one on  `myDocument` with one of two preset drops, depending on whether the custom drop value is greater than or less than half the height of the callout text box. For the example to work, shape one must be a callout.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(1).Callout

    If .DropType = msoCalloutDropCustom Then

        If .Drop < .Parent.Height / 2 Then

            .PresetDrop msoCalloutDropTop

        Else

            .PresetDrop msoCalloutDropBottom

        End If

    End If

End With
```


## See also


#### Concepts


[CalloutFormat Object](calloutformat-object-powerpoint.md)

