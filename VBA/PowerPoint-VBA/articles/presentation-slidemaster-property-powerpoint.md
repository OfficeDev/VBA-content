---
title: Presentation.SlideMaster Property (PowerPoint)
keywords: vbapp10.chm583003
f1_keywords:
- vbapp10.chm583003
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.SlideMaster
ms.assetid: 86b11fcd-b979-6ffe-bda7-1b9c6e807d29
ms.date: 06/08/2017
---


# Presentation.SlideMaster Property (PowerPoint)

Returns a  **[Master](master-object-powerpoint.md)** object that represents the slide master.


## Syntax

 _expression_. **SlideMaster**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

Master


## Example

This example sets the background pattern for the slide master for the active presentation.


```vb
Application.ActivePresentation.SlideMaster.Background.Fill _
    .PresetTextured msoTextureGreenMarble
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

