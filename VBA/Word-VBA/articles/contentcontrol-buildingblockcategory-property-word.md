---
title: ContentControl.BuildingBlockCategory Property (Word)
keywords: vbawd10.chm266534934
f1_keywords:
- vbawd10.chm266534934
ms.prod: word
api_name:
- Word.ContentControl.BuildingBlockCategory
ms.assetid: ca799bde-8556-381e-c9ca-74c5ac250d99
ms.date: 06/08/2017
---


# ContentControl.BuildingBlockCategory Property (Word)

Returns or sets a  **String** that represents the category for a building block content control. Read/write.


## Syntax

 _expression_ . **BuildingBlockCategory**

 _expression_ An expression that returns a **ContentControl** object.


## Remarks

This property applies only to building block content controls and corresponds with the  **Category** option in the **Content Control Properties** dialog box. You can set this property to any string; however, if you set it to a string for which there is no corresponding category, the value of the **Category** option is set to "(All Categories)".


## Example

The following example creates a new building block content control and specifies the type of building block and the gallery.


```vb
Dim objBB As ContentControl 
 
Set objBB = Selection.ContentControls.Add(wdContentControlBuildingBlockGallery) 
 
objBB.BuildingBlockType = wdTypeEquations 
objBB.BuildingBlockCategory = "General"
```


## See also


#### Concepts


[ContentControl Object](contentcontrol-object-word.md)

