---
title: Shape.Callout Property (PowerPoint)
keywords: vbapp10.chm547018
f1_keywords:
- vbapp10.chm547018
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Callout
ms.assetid: 381f8eaa-f373-b1aa-6a53-4086d7e887d8
ms.date: 06/08/2017
---


# Shape.Callout Property (PowerPoint)

Returns a  **[CalloutFormat](calloutformat-object-powerpoint.md)** object that contains callout formatting properties for the specified shape. Applies to **Shape** or **ShapeRange** objects that represent line callouts. Read-only.


## Syntax

 _expression_. **Callout**

 _expression_ A variable that represents a **Shape** object.


### Return Value

CalloutFormat


## Example

This example adds to  `myDocument` an oval and a callout that points to the oval. The callout text won't have a border, but it will have a vertical accent bar that separates the text from the callout line.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes

    .AddShape msoShapeOval, 180, 200, 280, 130

    With .AddCallout(msoCalloutTwo, 420, 170, 170, 40)

        .TextFrame.TextRange.Text = "My oval"

        With .Callout

            .Accent = True

            .Border = False

        End With

    End With

End With
```


## See also


#### Concepts


[Shape Object](shape-object-powerpoint.md)

