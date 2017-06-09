---
title: PageSetup.SlideSize Property (PowerPoint)
keywords: vbapp10.chm527006
f1_keywords:
- vbapp10.chm527006
ms.prod: powerpoint
api_name:
- PowerPoint.PageSetup.SlideSize
ms.assetid: 1f6db7f6-e9bb-e1fb-08f0-194b61733f5c
ms.date: 06/08/2017
---


# PageSetup.SlideSize Property (PowerPoint)

Returns or sets the slide size for the specified presentation. Read/write.


## Syntax

 _expression_. **SlideSize**

 _expression_ A variable that represents a **PageSetup** object.


### Return Value

PpSlideSizeType


## Remarks

The value of the  **SlideSize** property can be one of these **PpSlideSizeType** constants.


||
|:-----|
|**ppSlideSize35MM**|
|**ppSlideSizeA3Paper**|
|**ppSlideSizeA4Paper**|
|**ppSlideSizeB4ISOPaper**|
|**ppSlideSizeB4JISPaper**|
|**ppSlideSizeB5ISOPaper**|
|**ppSlideSizeB5JISPaper**|
|**ppSlideSizeBanner**|
|**ppSlideSizeCustom**|
|**ppSlideSizeHagakiCard**|
|**ppSlideSizeLedgerPaper**|
|**ppSlideSizeLetterPaper**|
|**ppSlideSizeOnScreen**|
|**ppSlideSizeOverhead**|

## Example

This example sets the slide size to overhead for the active presentation.


```vb
Application.ActivePresentation.PageSetup _
    .SlideSize = ppSlideSizeOverhead
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-powerpoint.md)

