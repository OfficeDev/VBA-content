---
title: CaptionLabel.NumberStyle Property (Word)
keywords: vbawd10.chm158924804
f1_keywords:
- vbawd10.chm158924804
ms.prod: word
api_name:
- Word.CaptionLabel.NumberStyle
ms.assetid: 1e668fdf-606c-04db-db3d-17284bd2d3af
ms.date: 06/08/2017
---


# CaptionLabel.NumberStyle Property (Word)

Returns or sets the number style for the  **CaptionLabel** object. Read/write **WdCaptionNumberStyle** .


## Syntax

 _expression_ . **NumberStyle**

 _expression_ Required. A variable that represents a **[CaptionLabel](captionlabel-object-word.md)** object.


## Remarks

Some of the constants listed above may not be available to you, depending on the language support (U.S. English, for example) that you've selected or installed.


## Example

This example inserts a caption at the insertion point. The caption letter is formatted as an uppercase letter.


```
CaptionLabels(wdCaptionFigure).NumberStyle = _ 
 wdCaptionNumberStyleUppercaseLetter 
Selection.Collapse Direction:=wdCollapseEnd 
Selection.InsertCaption Label:=wdCaptionFigure
```


## See also


#### Concepts


[CaptionLabel Object](captionlabel-object-word.md)

