---
title: Shape.ChangePicture Method (Visio)
ms.prod: VISIO
ms.assetid: 9193d802-cebd-2bfd-5f8e-400fac36c1a5
---


# Shape.ChangePicture Method (Visio)

Replaces the specified shape?s current picture with a new picture.


## Syntax

 _expression_ . **ChangePicture**_(FileName,_ _ChangePictureFlags)_

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|||||
| _FileName_|Required|STRING|Specifies the full path of the replacement picture.|
| _ChangePictureFlags_|Optional|INT32|Reserved for future implementation. Has no effect.|

### Return value

 **DOUBLE**


## Remarks

The  **DOUBLE** returned represents the ratio of the picture?s width to its height.


