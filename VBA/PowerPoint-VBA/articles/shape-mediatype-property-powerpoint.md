---
title: Shape.MediaType Property (PowerPoint)
keywords: vbapp10.chm547054
f1_keywords:
- vbapp10.chm547054
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.MediaType
ms.assetid: c42e3490-a4c9-d0bf-a201-71deff78d4b2
ms.date: 06/08/2017
---


# Shape.MediaType Property (PowerPoint)

Returns the OLE media type. Read-only.


## Syntax

 _expression_. **MediaType**

 _expression_ A variable that represents a **Shape** object.


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


[Shape Object](shape-object-powerpoint.md)

