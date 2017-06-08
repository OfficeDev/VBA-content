---
title: ShapeRange.MediaType Property (PowerPoint)
keywords: vbapp10.chm548054
f1_keywords:
- vbapp10.chm548054
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.MediaType
ms.assetid: 4d3d321c-6af5-36ce-5bf8-363dfce1a05f
ms.date: 06/08/2017
---


# ShapeRange.MediaType Property (PowerPoint)

Returns the OLE media type. Read-only.


## Syntax

 _expression_. **MediaType**

 _expression_ A variable that represents a **ShapeRange** object.


### Return Value

PpMediaType


## Remarks

The value of the  **MediaType** property can be one of these **PpMediaType** constants.


||
|:-----|
|**ppMediaTypeMixed**|
|**ppMediaTypeMovie**|
|**ppMediaTypeOther**|
|**ppMediaTypeSound**|

## Example

This example sets all native sound objects on slide one in the active presentation to loop until manually stopped during a slide show.


```vb
For Each so In ActivePresentation.Slides(1).Shapes

    If so.Type = msoMedia Then

        If so.MediaType = ppMediaTypeSound Then

            so.AnimationSettings.PlaySettings.LoopUntilStopped = True

        End If

    End If

Next
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)

