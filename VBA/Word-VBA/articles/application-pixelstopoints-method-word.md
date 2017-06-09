---
title: Application.PixelsToPoints Method (Word)
keywords: vbawd10.chm158335364
f1_keywords:
- vbawd10.chm158335364
ms.prod: word
api_name:
- Word.Application.PixelsToPoints
ms.assetid: f5e2e3f2-1e58-d84f-c73a-f6414fa48c3d
ms.date: 06/08/2017
---


# Application.PixelsToPoints Method (Word)

Converts a measurement from pixels to points. Returns the converted measurement as a  **Single** .


## Syntax

 _expression_ . **PixelsToPoints**( **_Pixels_** , **_fVertical_** )

 _expression_ Required. A variable that represents an **[Application](application-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Pixels_|Required| **Single**|The pixel value to be converted to points.|
| _fVertical_|Optional| **Variant**| **True** to convert vertical pixels; **False** to convert horizontal pixels.|

### Return Value

Single


## Example

This example displays the height and width in points of an object measured in pixels.


```vb
MsgBox "320x240 pixels is equivalent to " _ 
 &; PixelsToPoints(320, False) &; "x" _ 
 &; PixelsToPoints(240, True) _ 
 &; " points on this display."
```


## See also


#### Concepts


[Application Object](application-object-word.md)

