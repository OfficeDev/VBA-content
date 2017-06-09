---
title: CalloutFormat Object (PowerPoint)
keywords: vbapp10.chm559000
f1_keywords:
- vbapp10.chm559000
ms.prod: powerpoint
api_name:
- PowerPoint.CalloutFormat
ms.assetid: 7c06fe17-499e-b23c-3739-e53fe33d06f9
ms.date: 06/08/2017
---


# CalloutFormat Object (PowerPoint)

Contains properties and methods that apply to line callouts.


## Example

Use the  **Callout** property to return a **CalloutFormat** object. The following example specify the following attributes of shape three (a line callout) on `myDocument`:


- The callout will have a vertical accent bar that separates the text from the callout line.
    
- The angle between the callout line and the side of the callout text box will be 30 degrees.
    
- There will be no border around the callout text.
    
- The callout line will be attached to the top of the callout text box.
    
- The callout line will contain two segments.
    
For this example to work, shape three must be a callout.




```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3).Callout

    .Accent = True

    .Angle = msoCalloutAngle30

    .Border = False

    .PresetDrop msoCalloutDropTop

    .Type = msoCalloutThree

End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

