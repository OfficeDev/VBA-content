---
title: OLEFormat.FollowColors Property (PowerPoint)
keywords: vbapp10.chm562006
f1_keywords:
- vbapp10.chm562006
ms.prod: powerpoint
api_name:
- PowerPoint.OLEFormat.FollowColors
ms.assetid: 5f4c3f3d-0332-646f-de45-6854497f5782
ms.date: 06/08/2017
---


# OLEFormat.FollowColors Property (PowerPoint)

Returns or sets the extent to which the colors in the specified object follow the slide's color scheme. Read/write.


## Syntax

 _expression_. **FollowColors**

 _expression_ A variable that represents a **OLEFormat** object.


### Return Value

PpFollowColors


## Remarks

The specified object must be a chart created in either Microsoft Graph or Microsoft Organization Chart. 

The value of the  **FollowColors** property can be one of these **PpFollowColors** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**ppFollowColorsNone**| The chart colors don't follow the slide's color scheme.|
|**ppFollowColorsMixed**|Some of the chart colors follow the slide's color scheme.|
|**ppFollowColorsScheme**| All the colors in the chart follow the slide's color scheme.|
|**ppFollowColorsTextAndBackground**|Only the text and background follow the slide's color scheme.|

## Example

This example specifies that the text and background of shape two on slide one in the active presentation follow the slide's color scheme. Shape two must be a chart created in either Microsoft Graph or Microsoft Organization Chart.


```vb
ActivePresentation.Slides(1).Shapes(2).OLEFormat.FollowColors = ppFollowColorsTextAndBackground
```


## See also


#### Concepts


[OLEFormat Object](oleformat-object-powerpoint.md)

