---
title: Presentation.CreateVideo Method (PowerPoint)
keywords: vbapp10.chm583123
f1_keywords:
- vbapp10.chm583123
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.CreateVideo
ms.assetid: d302f251-66ee-c82d-d9b9-2c29b93f7615
ms.date: 06/08/2017
---


# Presentation.CreateVideo Method (PowerPoint)

Creates a video in a  **Presentation** object.


## Syntax

 _expression_. **CreateVideo**( **_FileName_**, **_UseTimingsAndNarrations_**, **_DefaultSlideDuration_**, **_VertResolution_**, **_FramesPerSecond_**, **_Quality_** )

 _expression_ A variable that represents a **Presentation** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The name of the video file to create.|
| _UseTimingsAndNarrations_|Optional|**Boolean**|Indicates whether to use timings and narrations.|
| _DefaultSlideDuration_|Optional|**[INT]**|The duration, in seconds, to view the slide.|
| _VertResolution_|Optional|**[INT]**|The resolution of the slide.|
| _FramesPerSecond_|Optional|**[INT]**|The number of frames per second.|
| _Quality_|Optional|**[INT]**|The level of quality of the slide.|

## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

