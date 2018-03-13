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
|<strong>ppSlideSize35MM</strong>|
|
<strong>ppSlideSizeA3Paper</strong>|
|
<strong>ppSlideSizeA4Paper</strong>|
|
<strong>ppSlideSizeB4ISOPaper</strong>|
|
<strong>ppSlideSizeB4JISPaper</strong>|
|
<strong>ppSlideSizeB5ISOPaper</strong>|
|
<strong>ppSlideSizeB5JISPaper</strong>|
|
<strong>ppSlideSizeBanner</strong>|
|
<strong>ppSlideSizeCustom</strong>|
|
<strong>ppSlideSizeHagakiCard</strong>|
|
<strong>ppSlideSizeLedgerPaper</strong>|
|
<strong>ppSlideSizeLetterPaper</strong>|
|
<strong>ppSlideSizeOnScreen</strong>|
|
<strong>ppSlideSizeOverhead</strong>|

## Example

This example sets the slide size to overhead for the active presentation.


```vb
Application.ActivePresentation.PageSetup _
    .SlideSize = ppSlideSizeOverhead
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-powerpoint.md)

