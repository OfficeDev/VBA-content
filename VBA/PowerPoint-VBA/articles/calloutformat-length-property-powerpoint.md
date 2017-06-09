---
title: CalloutFormat.Length Property (PowerPoint)
keywords: vbapp10.chm559014
f1_keywords:
- vbapp10.chm559014
ms.prod: powerpoint
api_name:
- PowerPoint.CalloutFormat.Length
ms.assetid: b0144e68-b495-0ef3-b228-599e56b7833e
ms.date: 06/08/2017
---


# CalloutFormat.Length Property (PowerPoint)

When the  **[AutoLength](calloutformat-autolength-property-powerpoint.md)** property of the specified callout is set to **False**, the **Length** property returns the length (in points) of the first segment of the callout line (the segment attached to the text callout box). Read-only.


## Syntax

 _expression_. **Length**

 _expression_ A variable that represents a **CalloutFormat** object.


## Remarks

Applies only to callouts whose lines consist of more than one segment (types  **msoCalloutThree** and **msoCalloutFour** ). Use the **[CustomLength](calloutformat-customlength-method-powerpoint.md)** method to set the value of this property for the **CalloutFormat** object.


## Example

If the first line segment in the callout named "co1" has a fixed length, this example specifies that the length of the first line segment in the callout named "co2" will also be fixed at that length. For the example to work, both callouts must have multiple-segment lines.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes

    With .Item("co1").Callout

        If Not .AutoLength Then len1 = .Length

    End With

    If len1 Then .Item("co2").Callout.CustomLength len1

End With
```


## See also


#### Concepts


[CalloutFormat Object](calloutformat-object-powerpoint.md)

