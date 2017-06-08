---
title: ListLevel Object (Word)
keywords: vbawd10.chm2445
f1_keywords:
- vbawd10.chm2445
ms.prod: word
api_name:
- Word.ListLevel
ms.assetid: 0cd152cb-6c25-50cb-7c1d-8b6d9734505b
ms.date: 06/08/2017
---


# ListLevel Object (Word)

Represents a single list level, either the only level for a bulleted or numbered list or one of the nine levels of an outline numbered list. The  **ListLevel** object is a member of the **ListLevels** collection.


## Remarks

Use  **ListLevels** (Index), where Index is a number from 1 through 9, to return a single **ListLevel** object. The following example sets list level one of list template one in the active document to start at 4.


```
ActiveDocument.ListTemplates(1).ListLevels(1).StartAt = 4
```

The  **ListLevel** object gives you access to all the formatting properties for the specified list level, such as the **Alignment**, **Font**, **NumberFormat**, **NumberPosition**, **NumberStyle**, and **TrailingCharacter** properties.

To apply a list level, first identify the range or list, and then use the  **ApplyListTemplate** method. Each tab at the beginning of the paragraph is translated into a list level. For example, a paragraph that begins with three tabs will become a level-three list paragraph after the **ApplyListTemplate** method is used.


## Methods



|**Name**|
|:-----|
|[ApplyPictureBullet](listlevel-applypicturebullet-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Alignment](listlevel-alignment-property-word.md)|
|[Application](listlevel-application-property-word.md)|
|[Creator](listlevel-creator-property-word.md)|
|[Font](listlevel-font-property-word.md)|
|[Index](listlevel-index-property-word.md)|
|[LinkedStyle](listlevel-linkedstyle-property-word.md)|
|[NumberFormat](listlevel-numberformat-property-word.md)|
|[NumberPosition](listlevel-numberposition-property-word.md)|
|[NumberStyle](listlevel-numberstyle-property-word.md)|
|[Parent](listlevel-parent-property-word.md)|
|[PictureBullet](listlevel-picturebullet-property-word.md)|
|[ResetOnHigher](listlevel-resetonhigher-property-word.md)|
|[StartAt](listlevel-startat-property-word.md)|
|[TabPosition](listlevel-tabposition-property-word.md)|
|[TextPosition](listlevel-textposition-property-word.md)|
|[TrailingCharacter](listlevel-trailingcharacter-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
