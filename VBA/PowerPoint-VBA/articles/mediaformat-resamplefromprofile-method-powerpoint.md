---
title: MediaFormat.ResampleFromProfile Method (PowerPoint)
keywords: vbapp10.chm724014
f1_keywords:
- vbapp10.chm724014
ms.prod: powerpoint
api_name:
- PowerPoint.MediaFormat.ResampleFromProfile
ms.assetid: f2d0ed29-82f1-e3f3-a4d9-e00a911176b3
ms.date: 06/08/2017
---


# MediaFormat.ResampleFromProfile Method (PowerPoint)

Adds the current media object to the queue and begins resampling base on the specified profile.


## Syntax

 _expression_. **ResampleFromProfile**( **_profile_** )

 _expression_ An expression that returns a **MediaFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _profile_|Optional|**PpResampleMediaProfile**|The resample media profile to use.|

### Return Value

Nothing


## Remarks

profile must be one of the following  **PpResampleMediaProfile** constants.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**ppResampleMediaProfileCustom**|1|Custom profile|
|**ppResampleMediaProfileSmall**|2|Small profile|
|**ppResampleMediaProfileSmaller**|3|Smaller profile|
|**ppResampleMediaProfileSmallest**|4|Smallest profile|

## See also


#### Concepts


[MediaFormat Object](mediaformat-object-powerpoint.md)

