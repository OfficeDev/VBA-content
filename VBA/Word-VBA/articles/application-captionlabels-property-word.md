---
title: Application.CaptionLabels Property (Word)
keywords: vbawd10.chm158334996
f1_keywords:
- vbawd10.chm158334996
ms.prod: word
api_name:
- Word.Application.CaptionLabels
ms.assetid: cf59346d-2ff5-938b-52ea-e2931422fd88
ms.date: 06/08/2017
---


# Application.CaptionLabels Property (Word)

Returns a  **[CaptionLabels](captionlabels-object-word.md)** collection that represents all the available caption labels. Read-only.


## Syntax

 _expression_ . **CaptionLabels**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example sets the numbering style for table captions.


```
CaptionLabels(wdCaptionTable).NumberStyle = _ 
 wdCaptionNumberStyleLowercaseRoman
```

This example adds a new caption label named "Photo" and then inserts a photo caption.




```vb
CaptionLabels.Add Name:="Photo" 
With Selection 
 .InsertParagraphAfter 
 .InsertCaption Label:="Photo" 
End With
```


## See also


#### Concepts


[Application Object](application-object-word.md)

