---
title: CaptionLabel Object (Word)
keywords: vbawd10.chm2425
f1_keywords:
- vbawd10.chm2425
ms.prod: word
api_name:
- Word.CaptionLabel
ms.assetid: 71c82dfd-6a66-e0f4-e30f-ae453c764864
ms.date: 06/08/2017
---


# CaptionLabel Object (Word)

Represents a single caption label. The  **CaptionLabel** object is a member of the **[CaptionLabels](captionlabels-object-word.md)** collection. The items in the **CaptionLabels** collection are listed in the **Label** box in the **Caption** dialog box.


## Remarks

Use  **[CaptionLabels](application-captionlabels-property-word.md)** (index), where index is the caption label name or index number, to return a single **CaptionLabel** object. The following example sets the numbering style for the Figure caption label.


```
CaptionLabels("Figure").NumberStyle = _ 
 wdCaptionNumberStyleLowercaseLetter
```

The index number represents the position of the caption label in the  **CaptionLabels** collection. The following example displays the first caption label.




```
MsgBox CaptionLabels(1).Name
```

Use the  **[Add](captionlabels-add-method-word.md)** method to add a custom caption label. The following example adds a caption label named "Photo."




```
CaptionLabels.Add Name:="Photo"
```


## Methods



|**Name**|
|:-----|
|[Delete](captionlabel-delete-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](captionlabel-application-property-word.md)|
|[BuiltIn](captionlabel-builtin-property-word.md)|
|[ChapterStyleLevel](captionlabel-chapterstylelevel-property-word.md)|
|[Creator](captionlabel-creator-property-word.md)|
|[ID](captionlabel-id-property-word.md)|
|[IncludeChapterNumber](captionlabel-includechapternumber-property-word.md)|
|[Name](captionlabel-name-property-word.md)|
|[NumberStyle](captionlabel-numberstyle-property-word.md)|
|[Parent](captionlabel-parent-property-word.md)|
|[Position](captionlabel-position-property-word.md)|
|[Separator](captionlabel-separator-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
