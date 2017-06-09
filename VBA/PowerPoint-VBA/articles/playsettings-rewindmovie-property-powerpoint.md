---
title: PlaySettings.RewindMovie Property (PowerPoint)
keywords: vbapp10.chm568007
f1_keywords:
- vbapp10.chm568007
ms.prod: powerpoint
api_name:
- PowerPoint.PlaySettings.RewindMovie
ms.assetid: 27eb1101-9604-e33c-1d7e-c8db643be1f9
ms.date: 06/08/2017
---


# PlaySettings.RewindMovie Property (PowerPoint)

Determines whether the first frame of the specified movie is automatically redisplayed as soon as the movie has finished playing. Read/write.


## Syntax

 _expression_. **RewindMovie**

 _expression_ A variable that represents a **PlaySettings** object.


### Return Value

MsoTriState


## Remarks

The value of the  **RewindMovie** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The first frame of the specified movie is not automatically redisplayed as soon as the movie has finished playing.|
|**msoTrue**| The first frame of the specified movie is automatically redisplayed as soon as the movie has finished playing.|

## Example

This example specifies that the first frame of the movie represented by shape three on slide one in the active presentation will be automatically redisplayed when the movie has finished playing. Shape three must be a movie object.


```vb
Set OLEobj = ActivePresentation.Slides(1).Shapes(3)

OLEobj.AnimationSettings.PlaySettings.RewindMovie = msoTrue
```


## See also


#### Concepts


[PlaySettings Object](playsettings-object-powerpoint.md)

