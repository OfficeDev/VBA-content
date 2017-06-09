---
title: PageSetup.SlideOrientation Property (PowerPoint)
keywords: vbapp10.chm527008
f1_keywords:
- vbapp10.chm527008
ms.prod: powerpoint
api_name:
- PowerPoint.PageSetup.SlideOrientation
ms.assetid: 24278d5b-075a-3f30-4667-b9c3af102382
ms.date: 06/08/2017
---


# PageSetup.SlideOrientation Property (PowerPoint)

Returns or sets the on-screen and printed orientation of slides in the specified presentation. Read/write.


## Syntax

 _expression_. **SlideOrientation**

 _expression_ A variable that represents a **PageSetup** object.


### Return Value

MsoOrientation


## Remarks

The value of the  **SlideOrientation** property can be one of these **MsoOrientation** constants.


||
|:-----|
|**msoOrientationHorizontal**|
|**msoOrientationMixed**|
|**msoOrientationVertical**|

## Example

This example sets orientation of all slides in the active presentation to vertical (portrait).


```vb
Application.ActivePresentation.PageSetup.SlideOrientation = msoOrientationVertical
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-powerpoint.md)

