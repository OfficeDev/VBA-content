---
title: Design.SlideMaster Property (PowerPoint)
keywords: vbapp10.chm644003
f1_keywords:
- vbapp10.chm644003
ms.prod: powerpoint
api_name:
- PowerPoint.Design.SlideMaster
ms.assetid: c6a9263c-462a-e9d8-7afc-32da3e133e90
ms.date: 06/08/2017
---


# Design.SlideMaster Property (PowerPoint)

Returns a  **[Master](master-object-powerpoint.md)** object that represents the slide master. Read-only.


## Syntax

 _expression_. **SlideMaster**

 _expression_ A variable that represents a **Design** object.


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


[Design Object](design-object-powerpoint.md)

