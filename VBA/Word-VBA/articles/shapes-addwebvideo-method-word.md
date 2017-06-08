---
title: Shapes.AddWebVideo Method (Word)
keywords: vbawd10.chm161415272
f1_keywords:
- vbawd10.chm161415272
ms.prod: word
ms.assetid: 9bdd1bc2-0d04-ca0c-eba2-4080843cf614
ms.date: 06/08/2017
---


# Shapes.AddWebVideo Method (Word)

Adds a new web video to the document.


## Syntax

 _expression_ . **AddWebVideo**_(EmbedCode,_ _VideoWidth,_ _VideoHeight,_ _PosterFrameImage,_ _Url,_ _Left,_ _Top,_ _Width,_ _Height,_ _Anchor)_

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|||||
| _EmbedCode_|Required|STRING|The HTML code to embed.|
| _VideoWidth_|Required|VARIANT|An integer that represents the width of the web video in pixels.|
| _VideoHeight_|Required|VARIANT|An integer that represents the height of the web video in pixels.|
| _PosterFrameImage_|Optional|VARIANT|A string that points to the file to use as the poster frame for the web video.|
| _Url_|Optional|VARIANT|A string that contains the URL to the web video.|
| _Left_|Optional|VARIANT|The position, measured in points, of the left edge of the poster frame from the edge of the document.|
| _Top_|Optional|VARIANT|The position, measured in points, of the top edge of the poster frame from the edge of the document.|
| _Width_|Optional|VARIANT|The width, measured in points, of the poster frame in the document.|
| _Height_|Optional|VARIANT|The height, measured in points, of the poster frame in the document.|
| _Anchor_|Optional|VARIANT|A [Range](range-object-word.md) object that represents the text to which the web video is bound. If _Anchor_ is specified, the anchor is positioned at the beginning of the first paragraph in the anchoring range. If this argument is omitted, the anchoring range is selected automatically and the video is positioned relative to the top and left edges of the page.|

### Return value

 **SHAPE**


## See also


#### Concepts


[Shapes Collection](shapes-object-word.md)

