---
title: CalloutFormat.Accent Property (PowerPoint)
keywords: vbapp10.chm559006
f1_keywords:
- vbapp10.chm559006
ms.prod: powerpoint
api_name:
- PowerPoint.CalloutFormat.Accent
ms.assetid: 901ad22d-2690-06c2-7327-9bf463585aa5
ms.date: 06/08/2017
---


# CalloutFormat.Accent Property (PowerPoint)

Determines whether a vertical accent bar separates the callout text from the callout line. Read/write.


## Syntax

 _expression_. **Accent**

 _expression_ A variable that represents an **CalloutFormat** object.


### Return Value

MsoTriState


## Remarks

The value of the  **Accent** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|A vertical accent bar does not separate the callout text from the callout line.|
|**msoTrue**| A vertical accent bar separates the callout text from the callout line.|

## Example

This example adds to  `myDocument` an oval and a callout that points to the oval. The callout text won't have a border, but it will have a vertical accent bar that separates the text from the callout line.


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

