---
title: CalloutFormat.AutomaticLength Method (PowerPoint)
keywords: vbapp10.chm559002
f1_keywords:
- vbapp10.chm559002
ms.prod: powerpoint
api_name:
- PowerPoint.CalloutFormat.AutomaticLength
ms.assetid: f80fdbbe-2fb4-c7d8-5f26-4edf16d65f82
ms.date: 06/08/2017
---


# CalloutFormat.AutomaticLength Method (PowerPoint)

Specifies that the first segment of the callout line (the segment attached to the text callout box) be scaled automatically when the callout is moved. Use the  **[CustomLength](calloutformat-customlength-method-powerpoint.md)** method to specify that the first segment of the callout line retain the fixed length returned by the **Length** property whenever the callout is moved. Applies only to callouts whose lines consist of more than one segment (types **msoCalloutThree** and **msoCalloutFour** ).


## Syntax

 _expression_. **AutomaticLength**

 _expression_ A variable that represents a **CalloutFormat** object.


## Remarks

Applying this method sets the [AutoLength](calloutformat-autolength-property-powerpoint.md)property to  **True**.


## Example

This example switches between an automatically scaling first segment and one with a fixed length for the callout line for shape one on  `myDocument`. For the example to work, shape one must be a callout.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(1).Callout

    If .AutoLength Then

        .CustomLength 50

    Else

        .AutomaticLength

    End If

End With
```


## See also


#### Concepts


[CalloutFormat Object](calloutformat-object-powerpoint.md)

