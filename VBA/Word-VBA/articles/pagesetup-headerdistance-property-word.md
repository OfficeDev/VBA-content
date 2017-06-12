---
title: PageSetup.HeaderDistance Property (Word)
keywords: vbawd10.chm158400624
f1_keywords:
- vbawd10.chm158400624
ms.prod: word
api_name:
- Word.PageSetup.HeaderDistance
ms.assetid: fee422f6-ecf0-0470-2845-b8694636a76e
ms.date: 06/08/2017
---


# PageSetup.HeaderDistance Property (Word)

Returns or sets the distance (in points) between the header and the top of the page. Read/write  **Single** .


## Syntax

 _expression_ . **HeaderDistance**

 _expression_ A variable that represents a **[PageSetup](pagesetup-object-word.md)** object.


## Example

This example displays the distance between the header and the top of the page. The  **[PointsToInches](global-pointstoinches-method-word.md)** method is used to convert points to inches.


```vb
Dim sngDistance As Single 
 
sngDistance = ActiveDocument.PageSetup.HeaderDistance 
Msgbox PointsToInches(sngDistance) &; " inches"
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-word.md)

