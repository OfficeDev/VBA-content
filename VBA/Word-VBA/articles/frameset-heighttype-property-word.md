---
title: Frameset.HeightType Property (Word)
keywords: vbawd10.chm165806082
f1_keywords:
- vbawd10.chm165806082
ms.prod: word
api_name:
- Word.Frameset.HeightType
ms.assetid: 4d83e41c-d33c-a5b8-853c-e7581170ba4b
ms.date: 06/08/2017
---


# Frameset.HeightType Property (Word)

Returns or sets the width type for the specified frame on a frames page. Read/write  **WdFramesetSizeType** .


## Syntax

 _expression_ . **HeightType**

 _expression_ Required. A variable that represents a **[Frameset](frameset-object-word.md)** object.


## Example

This example sets the height of the first Frameset object in the specified frames page to 25 percent of the window height.


```vb
With ActiveDocument.ActiveWindow.Panes(1).Frameset 
 .HeightType = wdFramesetSizeTypePercent 
 .Height = 25 
End With
```


## See also


#### Concepts


[Frameset Object](frameset-object-word.md)

