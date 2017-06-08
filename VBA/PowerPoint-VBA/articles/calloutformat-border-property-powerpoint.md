---
title: CalloutFormat.Border Property (PowerPoint)
keywords: vbapp10.chm559010
f1_keywords:
- vbapp10.chm559010
ms.prod: powerpoint
api_name:
- PowerPoint.CalloutFormat.Border
ms.assetid: 8183f14b-1432-300a-cf2b-650905661e53
ms.date: 06/08/2017
---


# CalloutFormat.Border Property (PowerPoint)

Determines whether the text in the specified callout is surrounded by a border. Read/write.


## Syntax

 _expression_. **Border**

 _expression_ A variable that represents a **CalloutFormat** object.


### Return Value

MsoTriState


## Remarks

The value of the  **Border** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The text in the specified callout is not surrounded by a border.|
|**msoTrue**| The text in the specified callout is surrounded by a border.|

## Example

This example adds to  `myDocument` an oval and a callout that points to the oval. The callout text does not have a border, but it does have a vertical accent bar that separates the text from the callout line.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes

    .AddShape msoShapeOval, 180, 200, 280, 130

    With .AddCallout(msoCalloutTwo, 420, 170, 170, 40)

        .TextFrame.TextRange.Text = "My oval"

        With .Callout

            .Accent = msoTrue

            .Border = msoFalse

        End With

    End With

End With
```


## See also


#### Concepts


[CalloutFormat Object](calloutformat-object-powerpoint.md)

