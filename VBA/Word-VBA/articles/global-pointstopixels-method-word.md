---
title: Global.PointsToPixels Method (Word)
keywords: vbawd10.chm163119489
f1_keywords:
- vbawd10.chm163119489
ms.prod: word
api_name:
- Word.Global.PointsToPixels
ms.assetid: e119ddf1-851c-2870-73f4-52da1d17c035
ms.date: 06/08/2017
---


# Global.PointsToPixels Method (Word)

Converts a measurement from points to pixels. Returns the converted measurement as a  **Single** .


## Syntax

 _expression_ . **PointsToPixels**( **_Points_** , **_fVertical_** )

 _expression_ Required. A variable that represents a **[Global](global-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Points_|Required| **Single**|The point value to be converted to pixels.|
| _fVertical_|Optional| **Variant**| **True** to return the result as vertical pixels; **False** to return the result as horizontal pixels.|

### Return Value

Single


## Example

This example displays the height and width in pixels of an object measured in points.


```vb
MsgBox "180x120 points is equivalent to " _ 
 &; PointsToPixels(180, False) &; "x" _ 
 &; PointsToPixels(120, True) _ 
 &; " pixels on this display."
```


## See also


#### Concepts


[Global Object](global-object-word.md)

