---
title: MediaFormat.ResamplingStatus Property (PowerPoint)
keywords: vbapp10.chm724015
f1_keywords:
- vbapp10.chm724015
ms.prod: powerpoint
api_name:
- PowerPoint.MediaFormat.ResamplingStatus
ms.assetid: 2a53f58e-3533-e93e-2aa1-9c6250f9c336
ms.date: 06/08/2017
---


# MediaFormat.ResamplingStatus Property (PowerPoint)

Returns the resampling task status. Read-only.


## Syntax

 _expression_. **ResamplingStatus**

 _expression_ An expression that returns a **MediaFormat** object.


### Return Value

PpMediaTaskStatus


## Remarks

 **ResamplingStatus** returns one of the following **PpMediaTaskStatus** values:



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**ppMediaTaskStatusNone**|0|No status|
|**ppMediaTaskStatusInProgress**|1|In progress|
|**ppMediaTaskStatusQueued**|2|Queued|
|**ppMediaTaskStatusDone**|3|Done|
|**ppMediaTaskStatusFailed**|4|Failed|

## See also


#### Concepts


[MediaFormat Object](mediaformat-object-powerpoint.md)

